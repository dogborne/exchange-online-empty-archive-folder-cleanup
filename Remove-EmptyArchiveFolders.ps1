<#
.SYNOPSIS
Safely reports and optionally deletes empty user-created folders from an Exchange Online archive mailbox by using EWS Managed API.

.DESCRIPTION
This script enumerates the complete archive hierarchy from ArchiveMsgFolderRoot by using paged, deep EWS FindFolders calls.
By default, it is report-only. Use -Delete to attempt deletion (deepest-first) with standard ShouldProcess semantics.

Prerequisites:
- Windows PowerShell 5.1+.
- EWS Managed API assembly (Microsoft.Exchange.WebServices.dll), supplied through -EwsManagedApiPath.
- MSAL.PS module with Get-MsalToken.
- Microsoft Entra app registration with Exchange Online application permission full_access_as_app, with admin consent.
- Certificate-based app-only auth certificate installed with private key in Cert:\CurrentUser\My or Cert:\LocalMachine\My.
- Recommended: scope access with an Exchange Application Access Policy.

Important:
- EWS retirement begins October 2026. This script is intended for one-off migration cleanup.
- Auto-expanding archive limitation: descendants beneath Deleted Items in auto-expanding archives are unsupported by design.
  The script does not perform automatic cross-archive routing.

.PARAMETER Mailbox
SMTP address of the target mailbox whose archive hierarchy will be inspected.

.PARAMETER TenantId
Microsoft Entra tenant ID (GUID or verified domain) used for OAuth token acquisition.

.PARAMETER ClientId
Application (client) ID of the Entra app registration configured for Exchange Online app-only EWS access.

.PARAMETER CertificateThumbprint
Thumbprint of the client certificate in Cert:\CurrentUser\My or Cert:\LocalMachine\My.

.PARAMETER EwsManagedApiPath
Path to Microsoft.Exchange.WebServices.dll, or a directory that contains it.

.PARAMETER Delete
When present, performs deletion attempts for validated candidates. Without this switch, the script only reports.

.PARAMETER LogPath
Optional CSV output path for UTF-8 logging. In report mode, logs candidates.
In delete mode, logs each attempted/skipped/succeeded/failed deletion action.

.EXAMPLE
.\Remove-EmptyArchiveFolders.ps1 `
  -Mailbox user@contoso.com `
  -TenantId contoso.onmicrosoft.com `
  -ClientId 11111111-2222-3333-4444-555555555555 `
  -CertificateThumbprint ABCDEF0123456789ABCDEF0123456789ABCDEF01 `
  -EwsManagedApiPath 'C:\Program Files\Microsoft\Exchange\Web Services\2.2' `
  -LogPath '.\ArchiveEmptyFolders-Report.csv'

Runs report-only mode and exports candidate folders to CSV.

.EXAMPLE
.\Remove-EmptyArchiveFolders.ps1 `
  -Mailbox user@contoso.com `
  -TenantId contoso.onmicrosoft.com `
  -ClientId 11111111-2222-3333-4444-555555555555 `
  -CertificateThumbprint ABCDEF0123456789ABCDEF0123456789ABCDEF01 `
  -EwsManagedApiPath 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll' `
  -Delete -WhatIf

Shows what would be deleted without deleting anything.

.EXAMPLE
.\Remove-EmptyArchiveFolders.ps1 `
  -Mailbox user@contoso.com `
  -TenantId contoso.onmicrosoft.com `
  -ClientId 11111111-2222-3333-4444-555555555555 `
  -CertificateThumbprint ABCDEF0123456789ABCDEF0123456789ABCDEF01 `
  -EwsManagedApiPath 'C:\Program Files\Microsoft\Exchange\Web Services\2.2' `
  -Delete -Confirm:$false -LogPath '.\ArchiveEmptyFolders-Delete.csv'

Performs actual hard deletions for validated candidates and writes a deletion log.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$EwsManagedApiPath,

    [Parameter()]
    [switch]$Delete,

    [Parameter()]
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$autoExpandingArchiveNotice = 'Descendants beneath Deleted Items in auto-expanding archives are unsupported by design; automatic cross-archive routing is not performed.'

function Resolve-EwsAssemblyPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath
    )

    try {
        $resolved = Resolve-Path -LiteralPath $InputPath -ErrorAction Stop | Select-Object -First 1
    }
    catch {
        throw "EwsManagedApiPath '$InputPath' was not found."
    }

    $providerPath = $resolved.ProviderPath
    $item = Get-Item -LiteralPath $providerPath -ErrorAction Stop

    if ($item.PSIsContainer) {
        $dllPath = Join-Path -Path $providerPath -ChildPath 'Microsoft.Exchange.WebServices.dll'
    }
    else {
        $dllPath = $providerPath
    }

    if (-not (Test-Path -LiteralPath $dllPath -PathType Leaf)) {
        throw "Microsoft.Exchange.WebServices.dll was not found at '$dllPath'."
    }

    return $dllPath
}

function Import-EwsManagedApiAssembly {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DllPath
    )

    if ('Microsoft.Exchange.WebServices.Data.ExchangeService' -as [type]) {
        return
    }

    try {
        Add-Type -Path $DllPath -ErrorAction Stop
    }
    catch {
        throw "Failed to load EWS Managed API assembly from '$DllPath'. Error: $($_.Exception.Message)"
    }
}

function Ensure-MsalPsAvailable {
    [CmdletBinding()]
    param()

    if (-not (Get-Module -ListAvailable -Name 'MSAL.PS')) {
        throw 'MSAL.PS module was not found. Install MSAL.PS before running this script.'
    }

    Import-Module -Name 'MSAL.PS' -ErrorAction Stop

    $msalCommand = Get-Command -Name 'Get-MsalToken' -Module 'MSAL.PS' -ErrorAction SilentlyContinue
    if (-not $msalCommand) {
        throw 'MSAL.PS is installed but Get-MsalToken is unavailable.'
    }
}

function Get-ClientCertificate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Thumbprint
    )

    $normalizedThumbprint = ($Thumbprint -replace '\s', '').ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($normalizedThumbprint)) {
        throw 'CertificateThumbprint cannot be empty.'
    }

    $certStores = @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My')
    $storeAccessErrors = @()
    foreach ($storePath in $certStores) {
        try {
            $cert = Get-ChildItem -Path $storePath -ErrorAction Stop |
                Where-Object { $_.Thumbprint -eq $normalizedThumbprint } |
                Select-Object -First 1
        }
        catch {
            $storeAccessErrors += "$storePath => $($_.Exception.Message)"
            continue
        }

        if ($cert) {
            if (-not $cert.HasPrivateKey) {
                throw "Certificate '$normalizedThumbprint' was found in '$storePath' but does not contain a private key."
            }

            if ($cert.NotAfter -lt (Get-Date)) {
                throw "Certificate '$normalizedThumbprint' is expired (NotAfter: $($cert.NotAfter.ToString('u')))."
            }

            return $cert
        }
    }

    if ($storeAccessErrors.Count -gt 0) {
        throw "Certificate '$normalizedThumbprint' was not found. Store access errors: $($storeAccessErrors -join ' | ')"
    }

    throw "Certificate '$normalizedThumbprint' was not found in Cert:\CurrentUser\My or Cert:\LocalMachine\My."
}

function Get-EwsErrorCode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Exception
    )

    if ($Exception -is [Microsoft.Exchange.WebServices.Data.ServiceResponseException]) {
        return [string]$Exception.ErrorCode
    }

    return ''
}

function Test-LikelyAutoExpandingArchiveCondition {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Exception
    )

    $errorCode = Get-EwsErrorCode -Exception $Exception
    $combined = '{0} {1}' -f $errorCode, $Exception.Message

    if ($combined -match '(?i)(redirect|moved|move|wrongserver|cross.?site|proxy|auxiliary|auto-?expanding|archive)') {
        return $true
    }

    $knownCodes = @(
        'ErrorCrossSiteRequest',
        'ErrorMailboxMoveInProgress',
        'ErrorWrongServerVersion',
        'ErrorProxyRequestNotAllowed',
        'ErrorNoRespondingCASInDestinationSite',
        'ErrorMailboxStoreUnavailable'
    )

    return ($knownCodes -contains $errorCode)
}

function Get-EwsAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantIdValue,

        [Parameter(Mandatory = $true)]
        [string]$ClientIdValue,

        [Parameter(Mandatory = $true)]
        $Certificate
    )

    try {
        $tokenResponse = Get-MsalToken `
            -TenantId $TenantIdValue `
            -ClientId $ClientIdValue `
            -ClientCertificate $Certificate `
            -Scopes 'https://outlook.office365.com/.default' `
            -ErrorAction Stop
    }
    catch {
        throw "Failed to acquire OAuth token from MSAL.PS. Error: $($_.Exception.Message)"
    }

    if (-not $tokenResponse.AccessToken) {
        throw 'MSAL.PS returned no access token.'
    }

    return $tokenResponse.AccessToken
}

function New-EwsService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $true)]
        [string]$MailboxSmtpAddress
    )

    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
    $service.Url = [Uri]'https://outlook.office365.com/EWS/Exchange.asmx'
    $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($AccessToken)
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
        [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,
        $MailboxSmtpAddress
    )
    $service.HttpHeaders['X-AnchorMailbox'] = $MailboxSmtpAddress
    $service.ReturnClientRequestId = $true
    $service.ClientRequestId = [Guid]::NewGuid().ToString()

    return $service
}

function New-FolderRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Folder,

        [Parameter(Mandatory = $true)]
        $AssocContentCountProperty,

        [Parameter(Mandatory = $true)]
        $AttrHiddenProperty
    )

    $assocCountValue = 0
    $associatedCountKnown = $Folder.TryGetProperty($AssocContentCountProperty, [ref]$assocCountValue)

    $hiddenValue = $false
    $hiddenKnown = $Folder.TryGetProperty($AttrHiddenProperty, [ref]$hiddenValue)

    return [pscustomobject]@{
        FolderId                  = $Folder.Id.UniqueId
        ParentFolderId            = if ($Folder.ParentFolderId) { $Folder.ParentFolderId.UniqueId } else { $null }
        DisplayName               = $Folder.DisplayName
        TotalCount                = [int]$Folder.TotalCount
        ChildFolderCount          = [int]$Folder.ChildFolderCount
        AssociatedCount           = if ($associatedCountKnown) { [int]$assocCountValue } else { 0 }
        AssociatedCountKnown      = [bool]$associatedCountKnown
        IsHidden                  = if ($hiddenKnown) { [bool]$hiddenValue } else { $false }
        IsHiddenKnown             = [bool]$hiddenKnown
        Depth                     = -1
        FolderPath                = $null
        IsWellKnownArchiveFolder  = $false
        IsUnderProtectedHierarchy = $false
        CanDelete                 = $false
        EligibilityReason         = ''
    }
}

function Get-FolderDepth {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderId,

        [Parameter(Mandatory = $true)]
        [hashtable]$Nodes,

        [Parameter(Mandatory = $true)]
        [hashtable]$DepthCache
    )

    if ($DepthCache.ContainsKey($FolderId)) {
        return [int]$DepthCache[$FolderId]
    }

    if (-not $Nodes.ContainsKey($FolderId)) {
        return 0
    }

    $node = $Nodes[$FolderId]
    if ([string]::IsNullOrWhiteSpace($node.ParentFolderId) -or -not $Nodes.ContainsKey($node.ParentFolderId)) {
        $depth = 0
    }
    else {
        $depth = (Get-FolderDepth -FolderId $node.ParentFolderId -Nodes $Nodes -DepthCache $DepthCache) + 1
    }

    $DepthCache[$FolderId] = $depth
    return $depth
}

function Get-FolderPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderId,

        [Parameter(Mandatory = $true)]
        [hashtable]$Nodes,

        [Parameter(Mandatory = $true)]
        [hashtable]$PathCache
    )

    if ($PathCache.ContainsKey($FolderId)) {
        return [string]$PathCache[$FolderId]
    }

    if (-not $Nodes.ContainsKey($FolderId)) {
        return '\<unknown>'
    }

    $node = $Nodes[$FolderId]
    $displayName = if ([string]::IsNullOrWhiteSpace($node.DisplayName)) { '<no name>' } else { $node.DisplayName }

    if ([string]::IsNullOrWhiteSpace($node.ParentFolderId) -or -not $Nodes.ContainsKey($node.ParentFolderId)) {
        $pathValue = '\' + $displayName
    }
    else {
        $parentPath = Get-FolderPath -FolderId $node.ParentFolderId -Nodes $Nodes -PathCache $PathCache
        if ($parentPath.EndsWith('\')) {
            $pathValue = $parentPath + $displayName
        }
        else {
            $pathValue = $parentPath + '\' + $displayName
        }
    }

    $PathCache[$FolderId] = $pathValue
    return $pathValue
}

function Test-IsUnderProtectedHierarchy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderId,

        [Parameter(Mandatory = $true)]
        [hashtable]$Nodes,

        [Parameter(Mandatory = $true)]
        $ProtectedSubtreeAnchors,

        [Parameter(Mandatory = $true)]
        [hashtable]$Cache
    )

    if ($Cache.ContainsKey($FolderId)) {
        return [bool]$Cache[$FolderId]
    }

    if (-not $Nodes.ContainsKey($FolderId)) {
        $Cache[$FolderId] = $false
        return $false
    }

    $node = $Nodes[$FolderId]
    $parentId = $node.ParentFolderId

    if ([string]::IsNullOrWhiteSpace($parentId)) {
        $Cache[$FolderId] = $false
        return $false
    }

    if ($ProtectedSubtreeAnchors.Contains($parentId)) {
        $Cache[$FolderId] = $true
        return $true
    }

    if (-not $Nodes.ContainsKey($parentId)) {
        $Cache[$FolderId] = $false
        return $false
    }

    $result = Test-IsUnderProtectedHierarchy -FolderId $parentId -Nodes $Nodes -ProtectedSubtreeAnchors $ProtectedSubtreeAnchors -Cache $Cache
    $Cache[$FolderId] = $result
    return $result
}

function New-ResultRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Mode,

        [Parameter(Mandatory = $true)]
        [string]$Action,

        [Parameter(Mandatory = $true)]
        [string]$Outcome,

        [Parameter(Mandatory = $true)]
        [string]$MailboxSmtpAddress,

        [Parameter()]
        $Node,

        [Parameter()]
        [string]$Reason = '',

        [Parameter()]
        [string]$ErrorCode = '',

        [Parameter()]
        [string]$ErrorMessage = ''
    )

    return [pscustomobject]@{
        TimestampUtc     = (Get-Date).ToUniversalTime().ToString('o')
        Mode             = $Mode
        Action           = $Action
        Outcome          = $Outcome
        Mailbox          = $MailboxSmtpAddress
        FolderPath       = if ($Node) { $Node.FolderPath } else { $null }
        FolderId         = if ($Node) { $Node.FolderId } else { $null }
        ParentFolderId   = if ($Node) { $Node.ParentFolderId } else { $null }
        DisplayName      = if ($Node) { $Node.DisplayName } else { $null }
        Depth            = if ($Node) { $Node.Depth } else { $null }
        TotalCount       = if ($Node) { $Node.TotalCount } else { $null }
        AssociatedCount  = if ($Node) { $Node.AssociatedCount } else { $null }
        ChildFolderCount = if ($Node) { $Node.ChildFolderCount } else { $null }
        Reason           = $Reason
        ErrorCode        = $ErrorCode
        ErrorMessage     = $ErrorMessage
    }
}

function Resolve-ArchiveWellKnownFolderIds {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Service,

        [Parameter(Mandatory = $true)]
        [string]$MailboxSmtpAddress,

        [Parameter(Mandatory = $true)]
        $PropertySet,

        [Parameter(Mandatory = $true)]
        [string]$AutoExpandingNoticeText
    )

    $enumType = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]
    $archiveWellKnownNames = [System.Enum]::GetNames($enumType) | Where-Object { $_ -like 'Archive*' }
    $resolvedIds = @{}
    $targetMailbox = New-Object Microsoft.Exchange.WebServices.Data.Mailbox($MailboxSmtpAddress)

    foreach ($name in $archiveWellKnownNames) {
        $enumValue = [System.Enum]::Parse($enumType, $name)
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($enumValue, $targetMailbox)

        try {
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $folderId, $PropertySet)
            if ($folder -and $folder.Id -and -not [string]::IsNullOrWhiteSpace($folder.Id.UniqueId)) {
                $resolvedIds[$name] = $folder.Id.UniqueId
            }
        }
        catch [Microsoft.Exchange.WebServices.Data.ServiceResponseException] {
            $errorCode = Get-EwsErrorCode -Exception $_.Exception
            if (Test-LikelyAutoExpandingArchiveCondition -Exception $_.Exception) {
                throw "Unable to resolve archive well-known folder '$name'. EWS error '$errorCode': $($_.Exception.Message). $AutoExpandingNoticeText"
            }

            if ($errorCode -eq 'ErrorFolderNotFound' -or $errorCode -eq 'ErrorItemNotFound') {
                continue
            }

            throw "Unable to resolve archive well-known folder '$name'. EWS error '$errorCode': $($_.Exception.Message)"
        }
        catch [Microsoft.Exchange.WebServices.Data.ServiceRequestException] {
            throw "Network/request error while resolving archive well-known folder '$name': $($_.Exception.Message)"
        }
    }

    return $resolvedIds
}

if ($Delete.IsPresent) {
    Write-Verbose 'Delete mode enabled.'
}
else {
    Write-Verbose 'Report-only mode (default).'
}

Write-Warning "Auto-expanding archive limitation: $autoExpandingArchiveNotice"

$dllPath = Resolve-EwsAssemblyPath -InputPath $EwsManagedApiPath
Import-EwsManagedApiAssembly -DllPath $dllPath
Ensure-MsalPsAvailable

$certificate = Get-ClientCertificate -Thumbprint $CertificateThumbprint
$accessToken = Get-EwsAccessToken -TenantIdValue $TenantId -ClientIdValue $ClientId -Certificate $certificate
$service = New-EwsService -AccessToken $accessToken -MailboxSmtpAddress $Mailbox

$assocContentCountProperty = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(
    0x3617,
    [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer
)
$attrHiddenProperty = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(
    0x10F4,
    [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean
)

$folderPropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
$null = $folderPropertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
$null = $folderPropertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId)
$null = $folderPropertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::ChildFolderCount)
$null = $folderPropertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::TotalCount)
$null = $folderPropertySet.Add($assocContentCountProperty)
$null = $folderPropertySet.Add($attrHiddenProperty)

$archiveMailbox = New-Object Microsoft.Exchange.WebServices.Data.Mailbox($Mailbox)
$archiveRootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
    [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot,
    $archiveMailbox
)

try {
    $archiveRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $archiveRootFolderId, $folderPropertySet)
}
catch [Microsoft.Exchange.WebServices.Data.ServiceResponseException] {
    $errorCode = Get-EwsErrorCode -Exception $_.Exception
    $message = "Failed to access archive root (ArchiveMsgFolderRoot) for '$Mailbox'. Ensure the mailbox has an archive and app access is allowed. EWS error '$errorCode': $($_.Exception.Message)"
    if (Test-LikelyAutoExpandingArchiveCondition -Exception $_.Exception) {
        $message = "$message $autoExpandingArchiveNotice"
    }

    throw $message
}
catch [Microsoft.Exchange.WebServices.Data.ServiceRequestException] {
    throw "Failed to access archive root due to request/network error: $($_.Exception.Message)"
}

$wellKnownFolderIds = Resolve-ArchiveWellKnownFolderIds -Service $service -MailboxSmtpAddress $Mailbox -PropertySet $folderPropertySet -AutoExpandingNoticeText $autoExpandingArchiveNotice
$wellKnownFolderIds['ArchiveMsgFolderRoot'] = $archiveRootFolder.Id.UniqueId

$requiredWellKnownNames = @(
    'ArchiveMsgFolderRoot',
    'ArchiveDeletedItems',
    'ArchiveRecoverableItemsRoot'
)
foreach ($requiredName in $requiredWellKnownNames) {
    if (-not $wellKnownFolderIds.ContainsKey($requiredName)) {
        throw "Safety check failed: required archive well-known folder '$requiredName' could not be resolved. This can indicate an unsupported auxiliary archive path. $autoExpandingArchiveNotice Aborting."
    }
}

$allFolders = New-Object 'System.Collections.Generic.List[object]'
$null = $allFolders.Add($archiveRootFolder)

$pageSize = 500
$offset = 0
$moreAvailable = $true

while ($moreAvailable) {
    $view = New-Object Microsoft.Exchange.WebServices.Data.FolderView(
        $pageSize,
        $offset,
        [Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning
    )
    $view.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    $view.PropertySet = $folderPropertySet

    try {
        $findResults = $service.FindFolders($archiveRootFolderId, $view)
    }
    catch [Microsoft.Exchange.WebServices.Data.ServiceResponseException] {
        $errorCode = Get-EwsErrorCode -Exception $_.Exception
        $message = "Failed during deep archive folder enumeration. EWS error '$errorCode': $($_.Exception.Message)"
        if (Test-LikelyAutoExpandingArchiveCondition -Exception $_.Exception) {
            $message = "$message $autoExpandingArchiveNotice"
        }

        throw $message
    }
    catch [Microsoft.Exchange.WebServices.Data.ServiceRequestException] {
        throw "Failed during deep archive folder enumeration due to request/network error: $($_.Exception.Message)"
    }

    foreach ($folder in $findResults.Folders) {
        $null = $allFolders.Add($folder)
    }

    $offset = $offset + $findResults.Folders.Count
    $moreAvailable = $findResults.MoreAvailable
}

$nodes = @{}
$childrenByParent = @{}

foreach ($folder in $allFolders) {
    $record = New-FolderRecord -Folder $folder -AssocContentCountProperty $assocContentCountProperty -AttrHiddenProperty $attrHiddenProperty
    $nodes[$record.FolderId] = $record
}

foreach ($node in $nodes.Values) {
    if (-not [string]::IsNullOrWhiteSpace($node.ParentFolderId)) {
        if (-not $childrenByParent.ContainsKey($node.ParentFolderId)) {
            $childrenByParent[$node.ParentFolderId] = New-Object 'System.Collections.Generic.List[string]'
        }

        $null = $childrenByParent[$node.ParentFolderId].Add($node.FolderId)
    }
}

$depthCache = @{}
$pathCache = @{}
foreach ($node in $nodes.Values) {
    $node.Depth = Get-FolderDepth -FolderId $node.FolderId -Nodes $nodes -DepthCache $depthCache
}
foreach ($node in $nodes.Values) {
    $node.FolderPath = Get-FolderPath -FolderId $node.FolderId -Nodes $nodes -PathCache $pathCache
}

$wellKnownById = @{}
foreach ($entry in $wellKnownFolderIds.GetEnumerator()) {
    if ($wellKnownById.ContainsKey($entry.Value)) {
        $wellKnownById[$entry.Value] = $wellKnownById[$entry.Value] + ',' + $entry.Key
    }
    else {
        $wellKnownById[$entry.Value] = $entry.Key
    }
}

$protectedWellKnownIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($id in $wellKnownFolderIds.Values) {
    if (-not [string]::IsNullOrWhiteSpace($id)) {
        $null = $protectedWellKnownIds.Add($id)
    }
}

$protectedSubtreeAnchors = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$null = $protectedSubtreeAnchors.Add($wellKnownFolderIds['ArchiveDeletedItems'])
foreach ($name in $wellKnownFolderIds.Keys) {
    if ($name -like 'ArchiveRecoverableItems*') {
        $null = $protectedSubtreeAnchors.Add($wellKnownFolderIds[$name])
    }
}

$underProtectedCache = @{}
$archiveRootId = $archiveRootFolder.Id.UniqueId
$nodesDescending = $nodes.Values | Sort-Object -Property @{ Expression = 'Depth'; Descending = $true }, @{ Expression = 'FolderPath'; Descending = $false }

foreach ($node in $nodesDescending) {
    $node.IsWellKnownArchiveFolder = $protectedWellKnownIds.Contains($node.FolderId)
    $node.IsUnderProtectedHierarchy = Test-IsUnderProtectedHierarchy -FolderId $node.FolderId -Nodes $nodes -ProtectedSubtreeAnchors $protectedSubtreeAnchors -Cache $underProtectedCache

    if ($node.FolderId -eq $archiveRootId) {
        $node.CanDelete = $false
        $node.EligibilityReason = 'Archive root is never deleted.'
        continue
    }

    if ($node.IsWellKnownArchiveFolder) {
        $wellKnownName = if ($wellKnownById.ContainsKey($node.FolderId)) { $wellKnownById[$node.FolderId] } else { 'ArchiveWellKnownFolder' }
        $node.CanDelete = $false
        $node.EligibilityReason = "Protected well-known archive folder ($wellKnownName)."
        continue
    }

    if ($node.IsUnderProtectedHierarchy) {
        $node.CanDelete = $false
        $node.EligibilityReason = 'Protected hierarchy: under Archive Deleted Items or Archive Recoverable Items.'
        continue
    }

    if (-not $node.IsHiddenKnown) {
        $node.CanDelete = $false
        $node.EligibilityReason = 'PR_ATTR_HIDDEN could not be read; skipping for safety.'
        continue
    }

    if (-not $node.AssociatedCountKnown) {
        $node.CanDelete = $false
        $node.EligibilityReason = 'PR_ASSOC_CONTENT_COUNT could not be read; skipping for safety.'
        continue
    }

    if ($node.IsHidden) {
        $node.CanDelete = $false
        $node.EligibilityReason = 'Folder is hidden (PR_ATTR_HIDDEN=True).'
        continue
    }

    if ($node.TotalCount -gt 0) {
        $node.CanDelete = $false
        $node.EligibilityReason = "Contains normal items (TotalCount=$($node.TotalCount))."
        continue
    }

    if ($node.AssociatedCount -gt 0) {
        $node.CanDelete = $false
        $node.EligibilityReason = "Contains associated/hidden items (PR_ASSOC_CONTENT_COUNT=$($node.AssociatedCount))."
        continue
    }

    $blockingChildren = @()
    if ($childrenByParent.ContainsKey($node.FolderId)) {
        foreach ($childId in $childrenByParent[$node.FolderId]) {
            $childNode = $nodes[$childId]
            if (-not $childNode.CanDelete) {
                $blockingChildren += $childId
            }
        }
    }

    if ($blockingChildren.Count -gt 0) {
        $node.CanDelete = $false
        $node.EligibilityReason = "Has non-deletable or nonempty descendants ($($blockingChildren.Count))."
        continue
    }

    $node.CanDelete = $true
    $node.EligibilityReason = 'Eligible empty user-created folder.'
}

$candidates = $nodes.Values |
    Where-Object { $_.CanDelete } |
    Sort-Object -Property @{ Expression = 'Depth'; Descending = $true }, @{ Expression = 'FolderPath'; Descending = $false }

$summary = [pscustomobject]@{
    RecordType                     = 'Summary'
    Mode                           = if ($Delete.IsPresent) { 'Delete' } else { 'Report' }
    Mailbox                        = $Mailbox
    TotalFoldersEnumerated         = $nodes.Count
    CandidateCount                 = $candidates.Count
    ProtectedWellKnownCount        = $protectedWellKnownIds.Count
    AutoExpandingArchiveLimitation = $autoExpandingArchiveNotice
    TimestampUtc                   = (Get-Date).ToUniversalTime().ToString('o')
}

$logRows = New-Object 'System.Collections.Generic.List[object]'
$outputRows = New-Object 'System.Collections.Generic.List[object]'
$null = $outputRows.Add($summary)

if (-not $Delete.IsPresent) {
    foreach ($candidate in $candidates) {
        $row = New-ResultRecord -Mode 'Report' -Action 'Candidate' -Outcome 'ReportOnly' -MailboxSmtpAddress $Mailbox -Node $candidate -Reason $candidate.EligibilityReason
        $null = $outputRows.Add($row)
        $null = $logRows.Add($row)
    }
}
else {
    foreach ($candidate in $candidates) {
        $actionDescription = "HardDelete archive folder '$($candidate.FolderPath)'"
        if (-not $PSCmdlet.ShouldProcess($candidate.FolderPath, $actionDescription)) {
            $row = New-ResultRecord -Mode 'Delete' -Action 'Delete' -Outcome 'SkippedShouldProcess' -MailboxSmtpAddress $Mailbox -Node $candidate -Reason 'Skipped by WhatIf or confirmation response.'
            $null = $outputRows.Add($row)
            $null = $logRows.Add($row)
            continue
        }

        $attemptRow = New-ResultRecord -Mode 'Delete' -Action 'Delete' -Outcome 'Attempted' -MailboxSmtpAddress $Mailbox -Node $candidate -Reason 'Deletion workflow started.'
        $null = $outputRows.Add($attemptRow)
        $null = $logRows.Add($attemptRow)

        $liveFolder = $null
        $liveAssocValue = 0
        $liveAssocKnown = $false
        $liveHiddenValue = $false
        $liveHiddenKnown = $false

        try {
            $liveFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($candidate.FolderId)
            $liveFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $liveFolderId, $folderPropertySet)
            $liveAssocKnown = $liveFolder.TryGetProperty($assocContentCountProperty, [ref]$liveAssocValue)
            $liveHiddenKnown = $liveFolder.TryGetProperty($attrHiddenProperty, [ref]$liveHiddenValue)
        }
        catch [Microsoft.Exchange.WebServices.Data.ServiceResponseException] {
            $errorCode = Get-EwsErrorCode -Exception $_.Exception
            $errorMessage = $_.Exception.Message
            if (Test-LikelyAutoExpandingArchiveCondition -Exception $_.Exception) {
                $errorMessage = "$errorMessage $autoExpandingArchiveNotice"
            }

            $row = New-ResultRecord -Mode 'Delete' -Action 'Rebind' -Outcome 'Failed' -MailboxSmtpAddress $Mailbox -Node $candidate -Reason 'Failed to rebind folder before delete.' -ErrorCode $errorCode -ErrorMessage $errorMessage
            $null = $outputRows.Add($row)
            $null = $logRows.Add($row)
            continue
        }
        catch [Microsoft.Exchange.WebServices.Data.ServiceRequestException] {
            $row = New-ResultRecord -Mode 'Delete' -Action 'Rebind' -Outcome 'Failed' -MailboxSmtpAddress $Mailbox -Node $candidate -Reason 'Request/network failure while rebinding folder before delete.' -ErrorMessage $_.Exception.Message
            $null = $outputRows.Add($row)
            $null = $logRows.Add($row)
            continue
        }

        $revalidationIssues = @()
        if ($liveFolder.TotalCount -ne 0) {
            $revalidationIssues += "TotalCount=$($liveFolder.TotalCount)"
        }
        if (-not $liveAssocKnown) {
            $revalidationIssues += 'PR_ASSOC_CONTENT_COUNT unavailable'
        }
        elseif ([int]$liveAssocValue -ne 0) {
            $revalidationIssues += "AssociatedCount=$([int]$liveAssocValue)"
        }
        if ($liveFolder.ChildFolderCount -ne 0) {
            $revalidationIssues += "ChildFolderCount=$($liveFolder.ChildFolderCount)"
        }
        if (-not $liveHiddenKnown) {
            $revalidationIssues += 'PR_ATTR_HIDDEN unavailable'
        }
        elseif ([bool]$liveHiddenValue) {
            $revalidationIssues += 'Hidden=True'
        }

        if ($revalidationIssues.Count -gt 0) {
            $reason = 'Revalidation failed immediately before delete: ' + ($revalidationIssues -join '; ')
            $row = New-ResultRecord -Mode 'Delete' -Action 'Revalidate' -Outcome 'Skipped' -MailboxSmtpAddress $Mailbox -Node $candidate -Reason $reason
            $null = $outputRows.Add($row)
            $null = $logRows.Add($row)
            continue
        }

        try {
            $liveFolder.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
            $row = New-ResultRecord -Mode 'Delete' -Action 'Delete' -Outcome 'Succeeded' -MailboxSmtpAddress $Mailbox -Node $candidate -Reason 'HardDelete succeeded.'
            $null = $outputRows.Add($row)
            $null = $logRows.Add($row)
        }
        catch [Microsoft.Exchange.WebServices.Data.ServiceResponseException] {
            $errorCode = Get-EwsErrorCode -Exception $_.Exception
            $errorMessage = $_.Exception.Message
            if (Test-LikelyAutoExpandingArchiveCondition -Exception $_.Exception) {
                $errorMessage = "$errorMessage $autoExpandingArchiveNotice"
            }

            $row = New-ResultRecord -Mode 'Delete' -Action 'Delete' -Outcome 'Failed' -MailboxSmtpAddress $Mailbox -Node $candidate -Reason 'EWS deletion failed.' -ErrorCode $errorCode -ErrorMessage $errorMessage
            $null = $outputRows.Add($row)
            $null = $logRows.Add($row)
            continue
        }
        catch [Microsoft.Exchange.WebServices.Data.ServiceRequestException] {
            $row = New-ResultRecord -Mode 'Delete' -Action 'Delete' -Outcome 'Failed' -MailboxSmtpAddress $Mailbox -Node $candidate -Reason 'Request/network failure during delete.' -ErrorMessage $_.Exception.Message
            $null = $outputRows.Add($row)
            $null = $logRows.Add($row)
            continue
        }
    }

    $deleteAttemptCount = ($logRows | Where-Object { $_.Mode -eq 'Delete' -and $_.Action -eq 'Delete' -and $_.Outcome -eq 'Attempted' }).Count
    $deleteSucceededCount = ($logRows | Where-Object { $_.Mode -eq 'Delete' -and $_.Action -eq 'Delete' -and $_.Outcome -eq 'Succeeded' }).Count
    $deleteSkippedCount = ($logRows | Where-Object { $_.Mode -eq 'Delete' -and $_.Outcome -like 'Skipped*' }).Count
    $deleteFailedCount = ($logRows | Where-Object { $_.Mode -eq 'Delete' -and $_.Outcome -eq 'Failed' }).Count

    $deleteSummary = [pscustomobject]@{
        RecordType      = 'DeleteSummary'
        Mailbox         = $Mailbox
        Candidates      = $candidates.Count
        DeleteAttempts  = $deleteAttemptCount
        DeleteSucceeded = $deleteSucceededCount
        DeleteSkipped   = $deleteSkippedCount
        DeleteFailed    = $deleteFailedCount
        TimestampUtc    = (Get-Date).ToUniversalTime().ToString('o')
    }

    $null = $outputRows.Add($deleteSummary)
}

if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
    $logDirectory = Split-Path -Path $LogPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($logDirectory) -and -not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
        throw "LogPath directory '$logDirectory' does not exist."
    }

    $logRows | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding UTF8
    $null = $outputRows.Add([pscustomobject]@{
            RecordType = 'Log'
            LogPath    = $LogPath
            Rows       = $logRows.Count
        })
}

$outputRows
