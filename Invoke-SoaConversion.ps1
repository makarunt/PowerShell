<#
.SYNOPSIS
    Inventory, convert, verify and roll back Source of Authority (SOA) for hybrid Microsoft 365 objects.

.DESCRIPTION
    Manages the Source of Authority (SOA) flag for objects synchronized from on-premises
    Active Directory with Microsoft Entra Connect Sync.

    One run handles exactly one Action and one Category:

      ExchangeAttributes  Set-Mailbox -IsExchangeCloudManaged (Exchange attributes of a mailbox)
      User                PATCH /users/{id}/onPremisesSyncBehavior    { "isCloudManaged": true|false }
      Group               PATCH /groups/{id}/onPremisesSyncBehavior   { "isCloudManaged": true|false }
      Contact             PATCH /contacts/{id}/onPremisesSyncBehavior { "isCloudManaged": true|false }

    The script only changes the SOA flag. It never changes other attributes, tenant-wide
    settings, on-premises AD, Entra Connect or Cloud Sync, and it never uses app-only
    authentication.

    Every run writes to a new folder OutputFolder\yyyyMMdd_HHmmss_<Action>_<Category>
    (transcript, inventory/backup CSV, results CSV). All CSV files are written and read as
    UTF-16 LE ("Unicode") with the delimiter from -Delimiter. ObjectId (Entra object ID,
    ExternalDirectoryObjectId for mailboxes) is the key in every CSV.

.PARAMETER Action
    Inventory, Convert, Verify or Rollback.

.PARAMETER Category
    ExchangeAttributes, User, Group or Contact.

.PARAMETER InputCsv
    CSV with an ObjectId column. Required for Convert, Verify and Rollback; optional for
    Inventory. For Convert and Rollback the CSV must also have a Category column whose value
    matches -Category on every row (an inventory or results CSV from this script works as-is).

.PARAMETER GroupType
    Only for -Category Group. Security, MailEnabledSecurity, Distribution or All (default).

.PARAMETER OutputFolder
    Root folder for run output. Default .\SoaOutput

.PARAMETER MaxObjects
    Process at most this many objects (first N rows of the CSV, or first N eligible objects
    for an Inventory without CSV). Use it for staged testing.

.PARAMETER Delimiter
    CSV delimiter for both import and export. Default ';'.

.PARAMETER ExpectedState
    Only for -Action Verify. CloudManaged (default) or OnPremManaged.

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Inventory -Category ExchangeAttributes
    Lists all dir-synced mailboxes whose Exchange attributes are still managed on-premises.

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Convert -Category ExchangeAttributes -InputCsv .\pilot.csv -MaxObjects 1
    Sets IsExchangeCloudManaged = $true for the first mailbox in pilot.csv, then verifies it.

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Verify -Category ExchangeAttributes -InputCsv .\pilot.csv
    Checks that the mailboxes in pilot.csv have IsExchangeCloudManaged = $true.

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Rollback -Category ExchangeAttributes -InputCsv .\pilot.csv -WhatIf
    Shows which mailboxes would be set back to IsExchangeCloudManaged = $false.

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Inventory -Category User -MaxObjects 50
    Lists the first 50 synced users that are not cloud managed yet.

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Convert -Category User -InputCsv .\users.csv -WhatIf
    Validates users.csv, takes a fresh backup and shows what would be converted. Nothing is changed.

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Verify -Category User -InputCsv .\users.csv
    Checks that isCloudManaged is true for every user in users.csv.

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Rollback -Category User -InputCsv .\users.csv
    Sets isCloudManaged back to false for the users in users.csv (see README for the tenant
    prerequisite blockCloudObjectTakeoverThroughHardMatchEnabled).

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Inventory -Category Group -GroupType Security
    Lists synced security groups (not mail-enabled), plus a separate group membership CSV.

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Convert -Category Group -GroupType Distribution -InputCsv .\dl.csv
    Converts the distribution groups in dl.csv. Groups of another type in the CSV are skipped.

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Verify -Category Group -InputCsv .\Results_Convert_Group.csv
    Re-checks the groups from an earlier results file.

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Rollback -Category Group -InputCsv .\groups.csv -Confirm
    Rolls back groups and asks for confirmation for every single group.

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Inventory -Category Contact -InputCsv .\contacts.csv
    Inventories only the contacts listed in contacts.csv and reports why others are not eligible.

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Convert -Category Contact -InputCsv .\contacts.csv -MaxObjects 5
    Converts the first five contacts from contacts.csv.

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Verify -Category Contact -InputCsv .\contacts.csv -ExpectedState OnPremManaged
    Checks that the contacts are no longer cloud managed (for example after a rollback).

.EXAMPLE
    .\Invoke-SoaConversion.ps1 -Action Rollback -Category Contact -InputCsv .\contacts.csv
    Sets isCloudManaged back to false for the contacts in contacts.csv.

.NOTES
    Requires PowerShell 7.2+, ExchangeOnlineManagement 3.2+ and Microsoft.Graph.Authentication 2.x.
    Windows PowerShell 5.1 is not supported.
    See README.md for roles, permissions and the recommended order of work.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Interactive console tool; output is captured by Start-Transcript.')]
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Inventory', 'Convert', 'Verify', 'Rollback')]
    [string]$Action,

    [Parameter(Mandatory)]
    [ValidateSet('ExchangeAttributes', 'User', 'Group', 'Contact')]
    [string]$Category,

    [string]$InputCsv,

    [ValidateSet('Security', 'MailEnabledSecurity', 'Distribution', 'All')]
    [string]$GroupType = 'All',

    [string]$OutputFolder = '.\SoaOutput',

    [ValidateRange(1, [int]::MaxValue)]
    [int]$MaxObjects,

    [ValidateNotNullOrEmpty()]
    [string]$Delimiter = ';',

    [ValidateSet('CloudManaged', 'OnPremManaged')]
    [string]$ExpectedState = 'CloudManaged'
)

# Checked here instead of with #Requires so that Windows PowerShell 5.1 gets a clear message.
# PowerShell parses the whole file before running it, so the whole file must stay parseable by 5.1
# (no ternary ?:, ??, ?., && or || operators) and ASCII-only (5.1 reads BOM-less files as ANSI).
$MinimumPowerShellVersion = [version]'7.2'
if ($PSVersionTable.PSVersion -lt $MinimumPowerShellVersion) {
    Write-Host "This script requires PowerShell $MinimumPowerShellVersion or newer (PowerShell 7)." -ForegroundColor Red
    Write-Host "Current version: $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))." -ForegroundColor Red
    Write-Host 'Install PowerShell 7:  winget install --id Microsoft.PowerShell --source winget' -ForegroundColor Yellow
    Write-Host '                  or:  https://aka.ms/powershell-release?tag=stable' -ForegroundColor Yellow
    Write-Host 'Then start "pwsh" (not "powershell") and run the script again.' -ForegroundColor Yellow
    exit 3
}

# ValidateSet accepts any casing ('user'); use the canonical names in file names and CSV columns.
$Action   = @('Inventory', 'Convert', 'Verify', 'Rollback') | Where-Object { $_ -eq $Action }
$Category = @('ExchangeAttributes', 'User', 'Group', 'Contact') | Where-Object { $_ -eq $Category }
$GroupType = @('Security', 'MailEnabledSecurity', 'Distribution', 'All') | Where-Object { $_ -eq $GroupType }
$ExpectedState = @('CloudManaged', 'OnPremManaged') | Where-Object { $_ -eq $ExpectedState }

#region Settings

$ErrorActionPreference = 'Stop'

$CsvEncoding         = 'Unicode'   # UTF-16 LE with BOM; keeps Croatian diacritics intact
$MultiValueSeparator = '|'
$VerifyAttempts      = 5
$VerifyDelaySeconds  = 20
$GraphRoot           = 'https://graph.microsoft.com'

# Microsoft Learn (Entra SOA how-to articles) uses v1.0. The Graph API reference documents the
# onPremisesSyncBehavior endpoint under beta only. The script starts with v1.0 and switches to
# beta for the SOA endpoint only if v1.0 rejects the onPremisesSyncBehavior segment.
$script:SoaApiVersion = 'v1.0'
$script:ExoConnected  = $false

$GraphSegment = @{ User = 'users'; Group = 'groups'; Contact = 'contacts' }

$GraphScopes = @{
    User    = @('User.Read.All', 'User-OnPremisesSyncBehavior.ReadWrite.All', 'RoleManagement.Read.Directory')
    Group   = @('Group.Read.All', 'GroupMember.Read.All', 'User.Read.All', 'Group-OnPremisesSyncBehavior.ReadWrite.All')
    Contact = @('OrgContact.Read.All', 'Contacts-OnPremisesSyncBehavior.ReadWrite.All')
}

$RequiredModules = @{
    'Microsoft.Graph.Authentication' = [version]'2.0.0'
    'ExchangeOnlineManagement'       = [version]'3.2.0'
}

$UserProperties = @(
    'userPrincipalName', 'displayName', 'givenName', 'surname', 'mail', 'mailNickname', 'proxyAddresses',
    'otherMails', 'jobTitle', 'department', 'companyName', 'officeLocation', 'employeeId', 'employeeType',
    'businessPhones', 'mobilePhone', 'faxNumber', 'streetAddress', 'city', 'state', 'postalCode', 'country',
    'usageLocation', 'accountEnabled', 'userType', 'onPremisesSyncEnabled', 'onPremisesImmutableId',
    'onPremisesDistinguishedName', 'onPremisesSamAccountName', 'onPremisesUserPrincipalName',
    'onPremisesDomainName', 'onPremisesSecurityIdentifier', 'onPremisesLastSyncDateTime'
)

$GroupProperties = @(
    'displayName', 'description', 'mail', 'mailNickname', 'proxyAddresses', 'securityEnabled', 'mailEnabled',
    'groupTypes', 'onPremisesSyncEnabled', 'onPremisesSamAccountName', 'onPremisesSecurityIdentifier',
    'onPremisesDomainName', 'onPremisesNetBiosName', 'onPremisesLastSyncDateTime'
)

$ContactProperties = @(
    'displayName', 'givenName', 'surname', 'mail', 'mailNickname', 'proxyAddresses', 'companyName',
    'department', 'jobTitle', 'phones', 'addresses', 'onPremisesSyncEnabled', 'onPremisesLastSyncDateTime'
)

$ExtensionAttributeNames = 1..15 | ForEach-Object { "extensionAttribute$_" }

$MailboxProperties = @(
    'DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress', 'WindowsEmailAddress', 'Alias', 'EmailAddresses',
    'RecipientTypeDetails', 'IsDirSynced', 'IsExchangeCloudManaged'
) + (1..15 | ForEach-Object { "CustomAttribute$_" }) + (1..5 | ForEach-Object { "ExtensionCustomAttribute$_" }) + @(
    'HiddenFromAddressListsEnabled', 'ForwardingAddress', 'ForwardingSmtpAddress', 'DeliverToMailboxAndForward',
    'GrantSendOnBehalfTo', 'AcceptMessagesOnlyFrom', 'AcceptMessagesOnlyFromDLMembers',
    'AcceptMessagesOnlyFromSendersOrMembers', 'RejectMessagesFrom', 'RejectMessagesFromDLMembers',
    'RejectMessagesFromSendersOrMembers', 'ModerationEnabled', 'ModeratedBy', 'RequireSenderAuthenticationEnabled',
    'LitigationHoldEnabled', 'RetentionHoldEnabled', 'ArchiveStatus', 'MailTip', 'Guid', 'ExchangeGuid',
    # Further attributes that Microsoft lists as editable in EXO after the Exchange-attribute SOA transfer
    'ArchiveGuid', 'ArchiveName', 'AuditEnabled', 'AuditLogAgeLimit', 'AuditAdmin', 'AuditDelegate', 'AuditOwner',
    'SingleItemRecoveryEnabled', 'LitigationHoldDate', 'LitigationHoldOwner', 'StartDateForRetentionHold',
    'EndDateForRetentionHold', 'RetentionComment', 'RetentionUrl', 'MailTipTranslations', 'SendModerationNotifications',
    'BypassModerationFromSendersOrMembers', 'ResourceCapacity', 'ResourceCustom', 'ImmutableId'
)

$ExoGroupProperties = @(
    'PrimarySmtpAddress', 'Alias', 'ManagedBy', 'HiddenFromAddressListsEnabled', 'RequireSenderAuthenticationEnabled',
    'AcceptMessagesOnlyFromSendersOrMembers', 'RejectMessagesFromSendersOrMembers', 'ModerationEnabled',
    'ModeratedBy', 'MemberJoinRestriction', 'MemberDepartRestriction', 'EmailAddresses'
)

$ExoContactProperties = @('ExternalEmailAddress', 'PrimarySmtpAddress', 'Alias', 'HiddenFromAddressListsEnabled', 'EmailAddresses')

# Protected system objects: never converted or rolled back, skipped with Note 'ProtectedSystemObject: ...'.
# A false match only skips an object, so the lists err on the side of caution. Edit them if needed.
# 1) Members of the Entra role 'Directory Synchronization Accounts' (Entra Connect / Cloud Sync service
#    accounts), whatever their name. Only relevant for Category User.
$DirSyncRoleTemplateId = 'd29b2b05-8046-44ba-8758-1e26182fcf32'
# 2) Wildcard patterns, matched case-insensitively against UPN, alias/mailNickname, display name and
#    on-premises SAM account name / UPN.
$ProtectedNamePatterns = @(
    'Sync_*'                                                  # Entra Connect sync account (Sync_<server>_<id>@...)
    'ADToAADSyncServiceAccount*'                              # Cloud Sync service account
    'On-Premises Directory Synchronization Service Account'   # display name of both
    'MSOL_*', 'AAD_*'                                         # Entra Connect AD DS connector accounts
    'Administrator', 'Guest', 'krbtgt', 'DefaultAccount', 'SUPPORT_388945a0*'
    'HealthMailbox*', 'SystemMailbox{*', 'DiscoverySearchMailbox*', 'FederatedEmail.*', 'Migration.8f3e7716*'
    'CAS_{*', 'SM_*', 'Exchange_Online-ApplicationAccount*'   # Exchange system mailboxes and accounts
)
# 3) Built-in Exchange and AD groups (matched exactly against group display name / SAM account name).
$ProtectedGroupNames = @(
    'Organization Management', 'View-Only Organization Management', 'Recipient Management', 'Server Management',
    'Discovery Management', 'Records Management', 'Compliance Management', 'Hygiene Management', 'Help Desk',
    'Public Folder Management', 'UM Management', 'Delegated Setup', 'Security Administrator', 'Security Reader',
    'Exchange Servers', 'Exchange Trusted Subsystem', 'Exchange Windows Permissions', 'ExchangeLegacyInterop',
    'Exchange Install Domain Servers', 'Managed Availability Servers', 'DnsAdmins', 'DnsUpdateProxy',
    'WinRMRemoteWMIUsers__'
)
# 4) Any user or group whose on-premises SID ends in a well-known RID (< 1000), e.g. Administrator (500),
#    krbtgt (502), Domain Admins (512).
$script:ProtectedObjectIds = [System.Collections.Generic.Dictionary[string, string]]::new([StringComparer]::OrdinalIgnoreCase)

#endregion

#region Console and file helpers

function Write-Log {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet('Info', 'Warn', 'Error', 'Success')][string]$Level = 'Info'
    )
    $color = @{ Info = 'Gray'; Warn = 'Yellow'; Error = 'Red'; Success = 'Green' }[$Level]
    Write-Host ('[{0:HH:mm:ss}] [{1,-7}] {2}' -f (Get-Date), $Level.ToUpper(), $Message) -ForegroundColor $color
}

function Show-WarningBlock {
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string[]]$Lines)
    $bar = '=' * 100
    Write-Host ''
    Write-Host $bar -ForegroundColor Yellow
    Write-Host "  WARNING: $Title" -ForegroundColor Yellow
    Write-Host $bar -ForegroundColor Yellow
    foreach ($line in $Lines) { Write-Host "  - $line" -ForegroundColor Yellow }
    Write-Host $bar -ForegroundColor Yellow
    Write-Host ''
}

function Confirm-TypedText {
    param([Parameter(Mandatory)][string]$Prompt, [string]$Expected = 'YES')
    $answer = Read-Host -Prompt "$Prompt Type '$Expected' to continue, anything else cancels"
    $accepted = $answer -ceq $Expected
    Write-Log "Operator answer: '$answer' -> $(if ($accepted) { 'accepted' } else { 'cancelled' })" -Level $(if ($accepted) { 'Info' } else { 'Warn' })
    return $accepted
}

function ConvertTo-CellValue {
    # Flattens a value into one CSV cell: multi-valued -> joined with '|', objects -> compact JSON.
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [datetime]) { return $Value.ToString('o') }
    if ($Value -is [ValueType]) { return [string]$Value }
    if ($Value -is [System.Collections.IDictionary]) { return ($Value | ConvertTo-Json -Compress -Depth 5) }
    if ($Value -is [System.Collections.IEnumerable]) {
        $items = foreach ($item in $Value) { ConvertTo-CellValue -Value $item }
        return (@($items) | Where-Object { $_ -ne '' }) -join $MultiValueSeparator
    }
    if ($Value.GetType().FullName -eq 'System.Management.Automation.PSCustomObject') {
        return ($Value | ConvertTo-Json -Compress -Depth 5)
    }
    return [string]$Value
}

function New-RunFolder {
    $root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
    $name = '{0:yyyyMMdd_HHmmss}_{1}_{2}' -f (Get-Date), $Action, $Category
    $path = Join-Path $root $name
    $suffix = 1
    while (Test-Path -LiteralPath $path) {
        $path = Join-Path $root ('{0}_{1}' -f $name, $suffix++)
    }
    New-Item -ItemType Directory -Path $path -Force -WhatIf:$false -Confirm:$false | Out-Null
    return $path
}

function Export-SoaCsv {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Data, [Parameter(Mandatory)][string]$Path)
    if ($Data.Count -eq 0) {
        Write-Log "No rows to write, skipped: $Path" -Level Warn
        return $null
    }
    if (Test-Path -LiteralPath $Path) { throw "Refusing to overwrite existing file: $Path" }
    $Data | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding $CsvEncoding -Delimiter $Delimiter -WhatIf:$false -Confirm:$false
    Write-Log "Written: $Path ($($Data.Count) rows)"
    return $Path
}

function Read-InputCsv {
    # Returns the de-duplicated, validated list of rows (ObjectId normalized to lower-case GUID).
    param([Parameter(Mandatory)][string]$Path, [switch]$RequireCategory)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Input CSV not found: $Path" }
    $fullPath = (Resolve-Path -LiteralPath $Path).ProviderPath

    $bom = [byte[]]::new(2)
    $stream = [System.IO.File]::OpenRead($fullPath)
    try { $null = $stream.Read($bom, 0, 2) } finally { $stream.Dispose() }
    if (-not ($bom[0] -eq 0xFF -and $bom[1] -eq 0xFE)) {
        throw "Input CSV '$Path' is not saved as UTF-16 LE (Unicode). Use a CSV created by this script, or delete rows in a text editor and save with 'UTF-16 LE' encoding."
    }

    $rows = @(Import-Csv -LiteralPath $fullPath -Encoding $CsvEncoding -Delimiter $Delimiter)
    if ($rows.Count -eq 0) { throw "Input CSV '$Path' has no data rows." }

    $columns = $rows[0].PSObject.Properties.Name
    if ('ObjectId' -notin $columns) {
        throw "Column 'ObjectId' not found in '$Path'. Columns read: [$($columns -join '], [')]. Check -Delimiter (current: '$Delimiter'). A file saved from Excel as 'Unicode Text' is tab-delimited: use -Delimiter `"``t`"."
    }

    if ('Category' -in $columns) {
        $mismatch = @($rows | Where-Object { "$($_.Category)".Trim() -ne $Category })
        if ($mismatch.Count -gt 0) {
            $found = ($mismatch | ForEach-Object { "'$("$($_.Category)".Trim())'" } | Sort-Object -Unique) -join ', '
            throw "Category mismatch: $($mismatch.Count) row(s) in '$Path' have Category $found, but -Category is '$Category'. Nothing was changed."
        }
    }
    elseif ($RequireCategory) {
        throw "Column 'Category' is required in '$Path' for -Action $Action. Use an inventory or results CSV from this script, or add a Category column with value '$Category'."
    }

    $invalid = [System.Collections.Generic.List[string]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $result = [System.Collections.Generic.List[object]]::new()
    $empty = 0; $duplicates = 0
    foreach ($row in $rows) {
        $raw = "$($row.ObjectId)".Trim()
        if (-not $raw) { $empty++; continue }
        $guid = [guid]::Empty
        if (-not [guid]::TryParse($raw, [ref]$guid)) { $invalid.Add($raw); continue }
        $id = $guid.ToString('D')
        if (-not $seen.Add($id)) { $duplicates++; continue }
        $row.ObjectId = $id
        $result.Add($row)
    }

    if ($invalid.Count -gt 0) {
        throw "Input CSV contains $($invalid.Count) invalid ObjectId value(s), for example: $(($invalid | Select-Object -First 5) -join ', '). Nothing was changed."
    }
    if ($empty -gt 0) { Write-Log "$empty row(s) with an empty ObjectId were ignored." -Level Warn }
    if ($duplicates -gt 0) { Write-Log "$duplicates duplicate ObjectId row(s) were ignored." -Level Warn }
    if ($result.Count -eq 0) { throw "Input CSV '$Path' contains no ObjectId values." }

    if ($MaxObjects -and $result.Count -gt $MaxObjects) {
        Write-Log "-MaxObjects ${MaxObjects}: using the first $MaxObjects of $($result.Count) rows." -Level Warn
        return @($result | Select-Object -First $MaxObjects)
    }
    return @($result)
}

#endregion

#region Error and Graph helpers

function Get-ErrorText {
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)
    $details = $ErrorRecord.ErrorDetails.Message
    if ($details) {
        try {
            $json = $details | ConvertFrom-Json -ErrorAction Stop
            if ($json.error.message) { return "$($json.error.code): $($json.error.message)" }
        }
        catch { Write-Verbose 'Error details are not JSON.' }
        return ($details -replace '\s+', ' ').Trim()
    }
    return ($ErrorRecord.Exception.Message -replace '\s+', ' ').Trim()
}

function Get-HttpStatusCode {
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)
    $response = $ErrorRecord.Exception.Response
    if ($response -and $response.StatusCode) { return [int]$response.StatusCode }

    $text = "$($ErrorRecord.Exception.Message) $($ErrorRecord.ErrorDetails.Message)"
    $names = [ordered]@{
        TooManyRequests = 429; ServiceUnavailable = 503; GatewayTimeout = 504; BadGateway = 502
        InternalServerError = 500; NotFound = 404; Forbidden = 403; Unauthorized = 401; BadRequest = 400
    }
    foreach ($name in $names.Keys) { if ($text -match "\b$name\b") { return $names[$name] } }
    if ($text -match '\b(400|401|403|404|429|500|502|503|504)\b') { return [int]$Matches[1] }
    return 0
}

function Get-RetryAfterSeconds {
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)
    try {
        $retryAfter = $ErrorRecord.Exception.Response.Headers.RetryAfter
        if ($retryAfter.Delta) { return [int][math]::Ceiling($retryAfter.Delta.TotalSeconds) }
        if ($retryAfter.Date) { return [int][math]::Max(1, [math]::Ceiling(($retryAfter.Date - [datetimeoffset]::UtcNow).TotalSeconds)) }
    }
    catch { Write-Verbose 'No Retry-After header available.' }
    return 0
}

function Invoke-GraphRequestWithRetry {
    # Retries on throttling (429) and transient server errors, honoring Retry-After.
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet('GET', 'PATCH')][string]$Method = 'GET',
        [hashtable]$Body,
        [hashtable]$Headers = @{},
        [int]$MaxAttempts = 6
    )
    for ($attempt = 1; ; $attempt++) {
        try {
            $request = @{ Method = $Method; Uri = $Uri; Headers = $Headers; OutputType = 'PSObject'; ErrorAction = 'Stop' }
            if ($Body) {
                $request.Body = $Body | ConvertTo-Json -Compress
                $request.ContentType = 'application/json'
            }
            return Invoke-MgGraphRequest @request
        }
        catch {
            $status = Get-HttpStatusCode -ErrorRecord $_
            if ($status -notin 429, 500, 502, 503, 504 -or $attempt -ge $MaxAttempts) { throw }
            $delay = Get-RetryAfterSeconds -ErrorRecord $_
            if ($delay -le 0) { $delay = [int][math]::Min(60, [math]::Pow(2, $attempt)) + (Get-Random -Maximum 3) }
            Write-Log "Graph returned HTTP $status for $Method (attempt $attempt/$MaxAttempts). Retrying in $delay s." -Level Warn
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-GraphCollection {
    # Follows @odata.nextLink. -Advanced adds ConsistencyLevel: eventual (needed for $count and
    # for $filter on onPremisesSyncEnabled).
    param([Parameter(Mandatory)][string]$Uri, [switch]$Advanced)
    $headers = @{}
    if ($Advanced) { $headers['ConsistencyLevel'] = 'eventual' }
    $next = $Uri
    while ($next) {
        $page = Invoke-GraphRequestWithRetry -Uri $next -Headers $headers
        foreach ($item in $page.value) { $item }
        $next = $page.'@odata.nextLink'
    }
}

function Get-GraphObjectById {
    param([Parameter(Mandatory)][string]$Segment, [Parameter(Mandatory)][string]$ObjectId, [Parameter(Mandatory)][string]$Select)
    try {
        return Invoke-GraphRequestWithRetry -Uri ("$GraphRoot/v1.0/$Segment/$ObjectId" + '?$select=' + $Select)
    }
    catch {
        if ((Get-HttpStatusCode -ErrorRecord $_) -eq 404) { return $null }
        throw
    }
}

function Get-GraphCloudManaged {
    param([Parameter(Mandatory)][string]$Segment, [Parameter(Mandatory)][string]$ObjectId)
    try {
        $behavior = Invoke-GraphRequestWithRetry -Uri "$GraphRoot/$script:SoaApiVersion/$Segment/$ObjectId/onPremisesSyncBehavior"
    }
    catch {
        if ($script:SoaApiVersion -eq 'v1.0' -and (Get-ErrorText -ErrorRecord $_) -match 'Resource not found for the segment') {
            Write-Log 'Graph v1.0 does not expose onPremisesSyncBehavior in this tenant. Switching the SOA endpoint to beta.' -Level Warn
            $script:SoaApiVersion = 'beta'
            return Get-GraphCloudManaged -Segment $Segment -ObjectId $ObjectId
        }
        throw
    }
    return ($behavior.isCloudManaged -eq $true)
}

function Get-GraphSoaStateName {
    # Microsoft documents: synced = (isCloudManaged false, onPremisesSyncEnabled true);
    # converted = (true, null); rolled back = (false, null) until the sync client takes over.
    param([AllowNull()][object]$GraphObject, [AllowNull()][object]$IsCloudManaged)
    if (-not $GraphObject) { return 'NotFound' }
    if ($IsCloudManaged -eq $true) { return 'CloudManaged' }
    if ($GraphObject.onPremisesSyncEnabled -eq $true) { return 'OnPremManaged' }
    return 'NotSynced'
}

#endregion

#region Prerequisites and connections

function Test-Prerequisites {
    param([bool]$NeedGraph, [bool]$NeedExo)
    Write-Log "PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
    $modules = @()
    if ($NeedExo) { $modules += 'ExchangeOnlineManagement' }
    if ($NeedGraph) { $modules += 'Microsoft.Graph.Authentication' }
    foreach ($name in $modules) {
        $minimum = $RequiredModules[$name]
        $loaded = Get-Module -Name $name | Sort-Object Version -Descending | Select-Object -First 1
        if ($loaded -and $loaded.Version -lt $minimum) {
            throw "Module '$name' $($loaded.Version) is already loaded in this session, but $minimum or newer is required. Close this PowerShell window and run the script in a new session."
        }
        $available = Get-Module -ListAvailable -Name $name | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $available -or $available.Version -lt $minimum) {
            Install-RequiredModule -Name $name -MinimumVersion $minimum -InstalledVersion $available.Version
            $available = Get-Module -ListAvailable -Name $name | Sort-Object Version -Descending | Select-Object -First 1
            if (-not $available -or $available.Version -lt $minimum) {
                throw "Module '$name' $minimum or newer is still not available after installation. Start a new PowerShell session and try again."
            }
        }
        Import-Module -Name $name -MinimumVersion $minimum -ErrorAction Stop
        Write-Log "Module $name $($available.Version) loaded."
    }
}

function Install-RequiredModule {
    # Offers to install (or update) a required module from the PowerShell Gallery for the current user.
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][version]$MinimumVersion, [version]$InstalledVersion)
    $manual = "Install-Module -Name $Name -MinimumVersion $MinimumVersion -Scope CurrentUser -Repository PSGallery"
    $problem = if ($InstalledVersion) { "version $InstalledVersion is installed, $MinimumVersion or newer is required" } else { 'it is not installed' }
    Write-Log "Module '$Name' is needed, but $problem." -Level Warn
    if (-not (Get-Command -Name Install-Module -ErrorAction SilentlyContinue)) {
        throw "Install-Module (PowerShellGet) is not available. Install the module manually: $manual"
    }
    $answer = Read-Host -Prompt "Install '$Name' from the PowerShell Gallery for the current user now? [Y/N]"
    if ($answer -notmatch '^\s*(y|yes)\s*$') {
        throw "Module '$Name' is required. Install it and run the script again: $manual"
    }
    Write-Log "Installing $Name (minimum $MinimumVersion) for the current user..."
    # -Force skips the 'untrusted repository' prompt (the operator just confirmed) and allows a side-by-side
    # install next to an older version. WhatIf/Confirm of the script do not apply to this local install.
    Install-Module -Name $Name -MinimumVersion $MinimumVersion -Scope CurrentUser -Repository PSGallery -Force -WhatIf:$false -Confirm:$false -ErrorAction Stop
    Write-Log "Module $Name installed." -Level Success
}

function Connect-SoaExchange {
    $WhatIfPreference = $false
    $connections = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' -and -not $_.IsEopSession })
    if ($connections.Count -eq 0) {
        Write-Log 'Connecting to Exchange Online (interactive sign-in)...'
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $connections = @(Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' -and -not $_.IsEopSession })
    }
    if ($connections.Count -ne 1) {
        throw "Expected exactly one Exchange Online connection, found $($connections.Count). Run Disconnect-ExchangeOnline and retry."
    }
    $connection = $connections[0]
    if ($connection.CertificateAuthentication -or [string]::IsNullOrWhiteSpace($connection.UserPrincipalName)) {
        throw 'The Exchange Online connection is app-only (certificate). Only delegated sign-in is supported. Run Disconnect-ExchangeOnline and retry.'
    }
    if ($connection.ModulePrefix) {
        throw "The Exchange Online connection uses the prefix '$($connection.ModulePrefix)'. Reconnect without -Prefix."
    }
    if ($Category -eq 'ExchangeAttributes') {
        $setMailbox = Get-Command -Name Set-Mailbox -ErrorAction SilentlyContinue
        if (-not $setMailbox -or -not $setMailbox.Parameters.ContainsKey('IsExchangeCloudManaged')) {
            throw 'Set-Mailbox has no -IsExchangeCloudManaged parameter in this session. Update ExchangeOnlineManagement and check that the feature is available in the tenant.'
        }
    }
    $script:ExoConnected = $true
    Write-Log "Exchange Online: $($connection.UserPrincipalName), tenant $($connection.TenantID)" -Level Success
    return $connection
}

function Connect-SoaGraph {
    param([Parameter(Mandatory)][string[]]$Scopes)
    $WhatIfPreference = $false
    $context = Get-MgContext
    if ($context -and $context.AuthType -ne 'Delegated') {
        throw "The Microsoft Graph connection uses '$($context.AuthType)' authentication. App-only is not allowed. Run Disconnect-MgGraph and retry."
    }
    $missing = if ($context) { @($Scopes | Where-Object { $_ -notin $context.Scopes }) } else { $Scopes }
    if ($missing.Count -gt 0) {
        Write-Log "Connecting to Microsoft Graph (interactive sign-in), scopes: $($Scopes -join ', ')"
        Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop
        $context = Get-MgContext
    }
    if ($context.AuthType -ne 'Delegated') { throw 'Microsoft Graph connection is not delegated.' }
    $missing = @($Scopes | Where-Object { $_ -notin $context.Scopes })
    if ($missing.Count -gt 0) {
        throw "Microsoft Graph scopes not granted: $($missing -join ', '). An Application Administrator or Cloud Application Administrator must consent to them."
    }
    Write-Log "Microsoft Graph: $($context.Account), tenant $($context.TenantId)" -Level Success
    return $context
}

#endregion

#region ExchangeAttributes (mailboxes)

function Get-MailboxByObjectId {
    param([Parameter(Mandatory)][string]$ObjectId)
    $found = @(Get-Mailbox -Filter "ExternalDirectoryObjectId -eq '$ObjectId'" -ResultSize 2 -WarningAction SilentlyContinue -ErrorAction Stop)
    if ($found.Count -gt 1) { throw "More than one mailbox has ExternalDirectoryObjectId $ObjectId." }
    if ($found.Count -eq 1) { return $found[0] }
    return $null
}

function Get-EligibleMailboxes {
    param([int]$Limit)
    $eligible = Get-Mailbox -ResultSize Unlimited -Filter 'IsDirSynced -eq $true' -WarningAction SilentlyContinue -ErrorAction Stop |
        Where-Object { $_.IsDirSynced -eq $true -and $_.IsExchangeCloudManaged -eq $false }
    if ($Limit) { $eligible = $eligible | Select-Object -First $Limit }
    return @($eligible)
}

function Get-MailboxSoaStateName {
    param([AllowNull()][object]$Mailbox)
    if (-not $Mailbox) { return 'NotFound' }
    if (-not $Mailbox.PSObject.Properties['IsExchangeCloudManaged']) { return 'Unknown' }
    if ($Mailbox.IsDirSynced -ne $true) { return 'NotSynced' }
    if ($Mailbox.IsExchangeCloudManaged -eq $true) { return 'CloudManaged' }
    return 'OnPremManaged'
}

function New-MailboxInventoryRow {
    param([Parameter(Mandatory)][string]$ObjectId, [AllowNull()][object]$Mailbox, [string]$LookupError)
    $columns = $MailboxProperties + @('HasUserCertificate', 'HasUserSMimeCertificate', 'AssistantName')
    $row = New-InventoryRow -ObjectId $ObjectId -Columns $columns
    if ($LookupError) { $row.Note = "Error: $LookupError"; return [pscustomobject]$row }

    $row.CurrentSoaState = Get-MailboxSoaStateName -Mailbox $Mailbox
    if ($Mailbox) {
        foreach ($name in $MailboxProperties) { $row[$name] = ConvertTo-CellValue $Mailbox.$name }
        $row.HasUserCertificate = (@($Mailbox.UserCertificate).Where({ $_ }).Count -gt 0)
        $row.HasUserSMimeCertificate = (@($Mailbox.UserSMimeCertificate).Where({ $_ }).Count -gt 0)
        # AssistantName (msExchAssistantName) is set with Set-User, so it is read with Get-User.
        try {
            $user = Get-User -Identity "$($Mailbox.Guid)" -ErrorAction Stop -WarningAction SilentlyContinue
            $row.AssistantName = ConvertTo-CellValue $user.AssistantName
        }
        catch { $row.Note = "Get-User failed, AssistantName not backed up: $(Get-ErrorText -ErrorRecord $_)" }
    }
    Set-RowEligibility -Row $row
    return [pscustomobject]$row
}

#endregion

#region User

function Get-EligibleUsers {
    param([int]$Limit)
    $select = (@('id') + $UserProperties + 'onPremisesExtensionAttributes') -join ','
    $uri = "$GraphRoot/v1.0/users?" + '$filter=' + [uri]::EscapeDataString('onPremisesSyncEnabled eq true') + '&$count=true&$top=999&$select=' + $select
    $users = Get-GraphCollection -Uri $uri -Advanced
    if ($Limit) { $users = $users | Select-Object -First $Limit }
    return @($users)
}

function New-UserInventoryRow {
    param([Parameter(Mandatory)][string]$ObjectId, [AllowNull()][object]$User)
    $columns = $UserProperties + 'managerUserPrincipalName' + $ExtensionAttributeNames + 'isCloudManaged'
    $row = New-InventoryRow -ObjectId $ObjectId -Columns $columns
    try {
        if (-not $User) {
            $select = (@('id') + $UserProperties + 'onPremisesExtensionAttributes') -join ','
            $User = Get-GraphObjectById -Segment 'users' -ObjectId $ObjectId -Select $select
        }
        if (-not $User) {
            $row.CurrentSoaState = 'NotFound'
            $row.Note = 'NotFound: no user with this ObjectId'
            return [pscustomobject]$row
        }
        foreach ($name in $UserProperties) { $row[$name] = ConvertTo-CellValue $User.$name }
        foreach ($name in $ExtensionAttributeNames) { $row[$name] = ConvertTo-CellValue $User.onPremisesExtensionAttributes.$name }

        # $expand instead of GET .../manager: a user without a manager then has no 'manager' property
        # instead of a 404, which Start-Transcript would log as an error even though it is handled.
        $managerUri = "$GraphRoot/v1.0/users/$ObjectId" + '?$select=id&$expand=manager($select=id,userPrincipalName,displayName)'
        $manager = (Invoke-GraphRequestWithRetry -Uri $managerUri).manager
        if ($manager) {
            $row.managerUserPrincipalName = ConvertTo-CellValue $(if ($manager.userPrincipalName) { $manager.userPrincipalName } else { $manager.displayName })
        }

        $isCloudManaged = Get-GraphCloudManaged -Segment 'users' -ObjectId $ObjectId
        $row.isCloudManaged = $isCloudManaged
        $row.CurrentSoaState = Get-GraphSoaStateName -GraphObject $User -IsCloudManaged $isCloudManaged
        Set-RowEligibility -Row $row
    }
    catch {
        $row.CurrentSoaState = 'Unknown'
        $row.Note = "Error: $(Get-ErrorText -ErrorRecord $_)"
    }
    return [pscustomobject]$row
}

#endregion

#region Group

function Get-GroupKind {
    param([Parameter(Mandatory)][object]$Group)
    if (@($Group.groupTypes) -contains 'Unified') { return 'Microsoft365' }
    if ($Group.securityEnabled -eq $true -and $Group.mailEnabled -ne $true) { return 'Security' }
    if ($Group.securityEnabled -eq $true -and $Group.mailEnabled -eq $true) { return 'MailEnabledSecurity' }
    if ($Group.securityEnabled -ne $true -and $Group.mailEnabled -eq $true) { return 'Distribution' }
    return 'Other'
}

function Get-SelectedGroupKinds {
    if ($GroupType -eq 'All') { return @('Security', 'MailEnabledSecurity', 'Distribution') }
    return @($GroupType)
}

function Get-EligibleGroups {
    param([int]$Limit)
    $select = (@('id') + $GroupProperties + 'onPremisesExtensionAttributes') -join ','
    $uri = "$GraphRoot/v1.0/groups?" + '$filter=' + [uri]::EscapeDataString('onPremisesSyncEnabled eq true') + '&$count=true&$top=999&$select=' + $select
    $kinds = Get-SelectedGroupKinds
    $groups = Get-GraphCollection -Uri $uri -Advanced | Where-Object { (Get-GroupKind -Group $_) -in $kinds }
    if ($Limit) { $groups = $groups | Select-Object -First $Limit }
    return @($groups)
}

function Get-ExoProperties {
    # Reads extra Exchange Online properties for a mail-enabled group or a mail contact.
    param([Parameter(Mandatory)][ValidateSet('Group', 'Contact')][string]$Kind, [Parameter(Mandatory)][string]$ObjectId)
    $filter = "ExternalDirectoryObjectId -eq '$ObjectId'"
    if ($Kind -eq 'Group') {
        $found = @(Get-DistributionGroup -Filter $filter -ResultSize 2 -WarningAction SilentlyContinue -ErrorAction Stop)
    }
    else {
        $found = @(Get-MailContact -Filter $filter -ResultSize 2 -WarningAction SilentlyContinue -ErrorAction Stop)
    }
    if ($found.Count -eq 1) { return $found[0] }
    return $null
}

function New-GroupInventoryRow {
    param(
        [Parameter(Mandatory)][string]$ObjectId,
        [AllowNull()][object]$Group,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Memberships
    )
    $exoColumns = $ExoGroupProperties | ForEach-Object { "Exo_$_" }
    $columns = @('GroupType') + $GroupProperties + $ExtensionAttributeNames + @(
        'isCloudManaged', 'MemberCount', 'CloudMemberCount', 'ChildGroupCount', 'ChildGroups', 'ParentGroupCount',
        'ParentGroups', 'IsNested', 'OwnerCount', 'Owners', 'SoaNotes'
    ) + $exoColumns
    $row = New-InventoryRow -ObjectId $ObjectId -Columns $columns
    try {
        if (-not $Group) {
            $select = (@('id') + $GroupProperties + 'onPremisesExtensionAttributes') -join ','
            $Group = Get-GraphObjectById -Segment 'groups' -ObjectId $ObjectId -Select $select
        }
        if (-not $Group) {
            $row.CurrentSoaState = 'NotFound'
            $row.Note = 'NotFound: no group with this ObjectId'
            return [pscustomobject]$row
        }
        $kind = Get-GroupKind -Group $Group
        $row.GroupType = $kind
        foreach ($name in $GroupProperties) { $row[$name] = ConvertTo-CellValue $Group.$name }
        foreach ($name in $ExtensionAttributeNames) { $row[$name] = ConvertTo-CellValue $Group.onPremisesExtensionAttributes.$name }

        $base = "$GraphRoot/v1.0/groups/$ObjectId"
        $members = @(Get-GraphCollection -Uri ("$base/members" + '?$select=id,displayName,userPrincipalName,mail,onPremisesSyncEnabled&$top=999'))
        $parents = @(Get-GraphCollection -Uri ("$base/memberOf" + '?$select=id,displayName&$top=999') | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' })
        $owners = @(Get-GraphCollection -Uri ("$base/owners" + '?$select=id,displayName,userPrincipalName&$top=999'))
        $childGroups = @($members | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' })

        $row.MemberCount = $members.Count
        $row.CloudMemberCount = @($members | Where-Object { $_.onPremisesSyncEnabled -ne $true }).Count
        $row.ChildGroupCount = $childGroups.Count
        $row.ChildGroups = ConvertTo-CellValue @($childGroups.displayName)
        $row.ParentGroupCount = $parents.Count
        $row.ParentGroups = ConvertTo-CellValue @($parents.displayName)
        $row.IsNested = ($childGroups.Count -gt 0 -or $parents.Count -gt 0)
        $row.OwnerCount = $owners.Count
        $row.Owners = ConvertTo-CellValue @($owners | ForEach-Object { if ($_.userPrincipalName) { $_.userPrincipalName } else { $_.displayName } })

        $isCloudManaged = Get-GraphCloudManaged -Segment 'groups' -ObjectId $ObjectId
        $row.isCloudManaged = $isCloudManaged
        $row.CurrentSoaState = Get-GraphSoaStateName -GraphObject $Group -IsCloudManaged $isCloudManaged

        foreach ($member in $members) {
            $Memberships.Add([pscustomobject][ordered]@{
                    GroupObjectId    = $ObjectId
                    GroupDisplayName = (ConvertTo-CellValue $Group.displayName)
                    MemberObjectId   = $member.id
                    MemberType       = ("$($member.'@odata.type')" -replace '^#microsoft\.graph\.', '')
                    MemberUPNorMail  = $(if ($member.userPrincipalName) { $member.userPrincipalName } else { "$($member.mail)" })
                    Category         = 'Group'
                    CurrentSoaState  = $row.CurrentSoaState
                })
        }

        $notes = [System.Collections.Generic.List[string]]::new()
        if ($row.IsNested) { $notes.Add('Nested group: SOA is not inherited by child or parent groups; convert each group separately, lowest level first.') }
        if ($kind -in 'MailEnabledSecurity', 'Distribution') {
            $notes.Add('Mail-enabled: after conversion manage only in Exchange Online (not in Entra admin center / Graph); no writeback to AD.')
            if ($script:ExoConnected) {
                $exoGroup = Get-ExoProperties -Kind Group -ObjectId $ObjectId
                if ($exoGroup) {
                    foreach ($name in $ExoGroupProperties) { $row["Exo_$name"] = ConvertTo-CellValue $exoGroup.$name }
                }
                else { $notes.Add('Exchange Online distribution group not found by ExternalDirectoryObjectId.') }
            }
        }
        if (@($ExtensionAttributeNames | Where-Object { $row[$_] }).Count -gt 0) {
            $notes.Add('Group has extensionAttributes 1-15 set; they are not supported after SOA conversion.')
        }
        $row.SoaNotes = $notes -join ' '

        Set-RowEligibility -Row $row
        if ($row.EligibleForConvert -and $kind -notin (Get-SelectedGroupKinds)) {
            $row.EligibleForConvert = $false
            $row.Note = "GroupTypeNotSelected: group type is $kind, -GroupType is $GroupType"
        }
        elseif ($kind -notin 'Security', 'MailEnabledSecurity', 'Distribution' -and $row.CurrentSoaState -eq 'OnPremManaged') {
            $row.EligibleForConvert = $false
            $row.Note = "UnsupportedGroupType: $kind"
        }
    }
    catch {
        $row.CurrentSoaState = 'Unknown'
        $row.Note = "Error: $(Get-ErrorText -ErrorRecord $_)"
    }
    return [pscustomobject]$row
}

#endregion

#region Contact

function Get-EligibleContacts {
    param([int]$Limit)
    $select = (@('id') + $ContactProperties) -join ','
    $uri = "$GraphRoot/v1.0/contacts?" + '$filter=' + [uri]::EscapeDataString('onPremisesSyncEnabled eq true') + '&$count=true&$top=100&$select=' + $select
    $contacts = Get-GraphCollection -Uri $uri -Advanced
    if ($Limit) { $contacts = $contacts | Select-Object -First $Limit }
    return @($contacts)
}

function New-ContactInventoryRow {
    param([Parameter(Mandatory)][string]$ObjectId, [AllowNull()][object]$Contact)
    $exoColumns = $ExoContactProperties | ForEach-Object { "Exo_$_" }
    $row = New-InventoryRow -ObjectId $ObjectId -Columns ($ContactProperties + 'isCloudManaged' + $exoColumns)
    try {
        if (-not $Contact) {
            $Contact = Get-GraphObjectById -Segment 'contacts' -ObjectId $ObjectId -Select ((@('id') + $ContactProperties) -join ',')
        }
        if (-not $Contact) {
            $row.CurrentSoaState = 'NotFound'
            $row.Note = 'NotFound: no contact with this ObjectId'
            return [pscustomobject]$row
        }
        foreach ($name in $ContactProperties) { $row[$name] = ConvertTo-CellValue $Contact.$name }
        $row.phones = ConvertTo-CellValue @($Contact.phones | Where-Object { $_.number } | ForEach-Object { "$($_.type):$($_.number)" })

        $isCloudManaged = Get-GraphCloudManaged -Segment 'contacts' -ObjectId $ObjectId
        $row.isCloudManaged = $isCloudManaged
        $row.CurrentSoaState = Get-GraphSoaStateName -GraphObject $Contact -IsCloudManaged $isCloudManaged

        if ($script:ExoConnected) {
            $exoContact = Get-ExoProperties -Kind Contact -ObjectId $ObjectId
            if ($exoContact) {
                foreach ($name in $ExoContactProperties) { $row["Exo_$name"] = ConvertTo-CellValue $exoContact.$name }
            }
        }
        Set-RowEligibility -Row $row
    }
    catch {
        $row.CurrentSoaState = 'Unknown'
        $row.Note = "Error: $(Get-ErrorText -ErrorRecord $_)"
    }
    return [pscustomobject]$row
}

#endregion

#region Protected system objects

function Initialize-ProtectedObjects {
    # Loads the members of 'Directory Synchronization Accounts'. Fails closed: if the role members cannot be
    # read, the run stops instead of risking a change to a sync service account.
    $uri = "$GraphRoot/v1.0/directoryRoles(roleTemplateId='$DirSyncRoleTemplateId')/members" + '?$select=id,userPrincipalName,displayName'
    try {
        $members = @(Get-GraphCollection -Uri $uri)
    }
    catch {
        if ((Get-HttpStatusCode -ErrorRecord $_) -eq 404) {
            $members = @()   # role not activated in the tenant: no members
        }
        else {
            throw "Could not read members of the 'Directory Synchronization Accounts' role, so sync service accounts cannot be protected: $(Get-ErrorText -ErrorRecord $_)"
        }
    }
    foreach ($member in $members) {
        $script:ProtectedObjectIds[$member.id] = "member of 'Directory Synchronization Accounts' role"
        Write-Log "Protected sync service account: $($member.userPrincipalName) ($($member.id))"
    }
}

function Get-ProtectedReason {
    param([Parameter(Mandatory)][object]$Row)
    $properties = $Row.PSObject.Properties
    if ($script:ProtectedObjectIds.ContainsKey("$($Row.ObjectId)")) { return $script:ProtectedObjectIds["$($Row.ObjectId)"] }

    foreach ($field in 'userPrincipalName', 'mailNickname', 'Alias', 'displayName', 'onPremisesSamAccountName', 'onPremisesUserPrincipalName') {
        if (-not $properties[$field]) { continue }
        $value = "$($properties[$field].Value)"
        if (-not $value) { continue }
        foreach ($pattern in $ProtectedNamePatterns) {
            if ($value -like $pattern) { return "$field '$value' matches system account pattern '$pattern'" }
        }
        if ($Category -eq 'Group' -and $field -in 'displayName', 'onPremisesSamAccountName' -and $value -in $ProtectedGroupNames) {
            return "built-in group '$value'"
        }
    }

    if ($properties['onPremisesSecurityIdentifier']) {
        $sid = "$($properties['onPremisesSecurityIdentifier'].Value)"
        if ($sid -match '-(\d+)$' -and [long]$Matches[1] -lt 1000) { return "well-known on-premises SID $sid" }
    }
    return $null
}

function Set-RowProtection {
    param([Parameter(Mandatory)][object]$Row)
    $reason = Get-ProtectedReason -Row $Row
    if ($reason) {
        $Row.EligibleForConvert = $false
        $Row.Note = "ProtectedSystemObject: $reason"
    }
}

function Test-ProtectedRow {
    param([Parameter(Mandatory)][object]$Row)
    return "$($Row.Note)" -like 'ProtectedSystemObject:*'
}

#endregion

#region Inventory, state, change

function New-InventoryRow {
    param([Parameter(Mandatory)][string]$ObjectId, [Parameter(Mandatory)][string[]]$Columns)
    $row = [ordered]@{
        ObjectId           = $ObjectId
        Category           = $Category
        CurrentSoaState    = 'Unknown'
        EligibleForConvert = $false
        Note               = ''
    }
    foreach ($column in $Columns) { $row[$column] = '' }
    return $row
}

function Set-RowEligibility {
    param([Parameter(Mandatory)][System.Collections.Specialized.OrderedDictionary]$Row)
    switch ($Row.CurrentSoaState) {
        'OnPremManaged' { $Row.EligibleForConvert = $true }
        'CloudManaged'  { $Row.Note = 'AlreadyCloudManaged' }
        'NotSynced'     { $Row.Note = 'NotSynced: not synchronized from AD (cloud-only object, or a rollback waiting for the next sync cycle)' }
        'NotFound'      { $Row.Note = 'NotFound' }
        default         { $Row.Note = "UnknownState: $($Row.CurrentSoaState)" }
    }
}

function Invoke-SoaInventory {
    # Returns @{ Rows; Memberships }. With -ObjectIds: one row per ID (including NotFound).
    # Without: only eligible objects.
    param([string[]]$ObjectIds)
    $rows = [System.Collections.Generic.List[object]]::new()
    $memberships = [System.Collections.Generic.List[object]]::new()
    $fromCsv = $null -ne $ObjectIds

    if ($fromCsv) {
        $items = $ObjectIds | ForEach-Object { [pscustomobject]@{ Id = $_; Object = $null } }
    }
    else {
        Write-Log "Querying all eligible $Category objects..."
        $objects = switch ($Category) {
            'ExchangeAttributes' { Get-EligibleMailboxes -Limit $MaxObjects }
            'User'               { Get-EligibleUsers -Limit $MaxObjects }
            'Group'              { Get-EligibleGroups -Limit $MaxObjects }
            'Contact'            { Get-EligibleContacts -Limit $MaxObjects }
        }
        $items = $objects | ForEach-Object {
            $id = if ($Category -eq 'ExchangeAttributes') { "$($_.ExternalDirectoryObjectId)" } else { "$($_.id)" }
            [pscustomobject]@{ Id = $id; Object = $_ }
        }
    }
    $items = @($items)
    Write-Log "Reading details for $($items.Count) object(s)..."

    $index = 0
    foreach ($item in $items) {
        $index++
        Write-Progress -Activity "Inventory $Category" -Status "$index / $($items.Count)" -PercentComplete ($index * 100 / [math]::Max(1, $items.Count))
        switch ($Category) {
            'ExchangeAttributes' {
                $lookupError = $null
                $mailbox = $item.Object
                if (-not $mailbox) {
                    try { $mailbox = Get-MailboxByObjectId -ObjectId $item.Id }
                    catch { $lookupError = Get-ErrorText -ErrorRecord $_ }
                }
                $rows.Add((New-MailboxInventoryRow -ObjectId $item.Id -Mailbox $mailbox -LookupError $lookupError))
            }
            'User'    { $rows.Add((New-UserInventoryRow -ObjectId $item.Id -User $item.Object)) }
            'Group'   { $rows.Add((New-GroupInventoryRow -ObjectId $item.Id -Group $item.Object -Memberships $memberships)) }
            'Contact' { $rows.Add((New-ContactInventoryRow -ObjectId $item.Id -Contact $item.Object)) }
        }
        Set-RowProtection -Row $rows[$rows.Count - 1]
    }
    Write-Progress -Activity "Inventory $Category" -Completed

    $protectedCount = @($rows | Where-Object { Test-ProtectedRow -Row $_ }).Count
    if ($protectedCount -gt 0) {
        Write-Log "$protectedCount protected system object(s) found; they are never changed by this script (see Note column)." -Level Warn
    }

    if (-not $fromCsv) {
        # Keep eligible objects and objects that could not be read (Error), so read failures stay visible.
        $leftOut = @($rows | Where-Object { -not $_.EligibleForConvert -and "$($_.Note)" -notlike 'Error:*' })
        if ($leftOut.Count -gt 0) {
            Write-Log "$($leftOut.Count) listed object(s) are not eligible and were left out:" -Level Warn
            foreach ($group in ($leftOut | Group-Object { ($_.Note -split ':')[0] })) { Write-Log "  $($group.Count) x $($group.Name)" -Level Warn }
            $keep = [System.Collections.Generic.HashSet[string]]::new([string[]]@($rows | Where-Object { $_ -notin $leftOut } | ForEach-Object ObjectId))
            $rows = [System.Collections.Generic.List[object]]@($rows | Where-Object { $keep.Contains($_.ObjectId) })
            $memberships = [System.Collections.Generic.List[object]]@($memberships | Where-Object { $keep.Contains($_.GroupObjectId) })
        }
    }
    $errors = @($rows | Where-Object { "$($_.Note)" -like 'Error:*' })
    if ($errors.Count -gt 0) {
        Write-Log "$($errors.Count) object(s) could not be read completely (Note starts with 'Error:'); they are never changed." -Level Error
    }
    return @{ Rows = $rows; Memberships = $memberships }
}

function Get-ObjectSoaState {
    # Current SOA state of one object: @{ State; DisplayName; Error }.
    param([Parameter(Mandatory)][string]$ObjectId)
    try {
        if ($Category -eq 'ExchangeAttributes') {
            $mailbox = Get-MailboxByObjectId -ObjectId $ObjectId
            return [pscustomobject]@{ State = (Get-MailboxSoaStateName -Mailbox $mailbox); DisplayName = "$($mailbox.DisplayName)"; Error = '' }
        }
        $segment = $GraphSegment[$Category]
        $object = Get-GraphObjectById -Segment $segment -ObjectId $ObjectId -Select 'id,displayName,onPremisesSyncEnabled'
        $isCloudManaged = if ($object) { Get-GraphCloudManaged -Segment $segment -ObjectId $ObjectId } else { $null }
        return [pscustomobject]@{ State = (Get-GraphSoaStateName -GraphObject $object -IsCloudManaged $isCloudManaged); DisplayName = "$($object.displayName)"; Error = '' }
    }
    catch {
        return [pscustomobject]@{ State = 'Unknown'; DisplayName = ''; Error = (Get-ErrorText -ErrorRecord $_) }
    }
}

function Test-StateMatchesExpected {
    param([string]$State, [ValidateSet('CloudManaged', 'OnPremManaged')][string]$Expected)
    if ($Expected -eq 'CloudManaged') { return $State -eq 'CloudManaged' }
    # After a Graph rollback the object reads isCloudManaged=false and onPremisesSyncEnabled=null
    # until Connect Sync takes it over; that is the expected intermediate state.
    if ($Category -eq 'ExchangeAttributes') { return $State -eq 'OnPremManaged' }
    return $State -in 'OnPremManaged', 'NotSynced'
}

function Set-ObjectSoa {
    # Changes only the SOA flag. Nothing else is sent to the service.
    param([Parameter(Mandatory)][string]$ObjectId, [Parameter(Mandatory)][bool]$CloudManaged)
    if ($Category -eq 'ExchangeAttributes') {
        # IsExchangeCloudManaged has its own parameter set; it must be the only parameter.
        Set-Mailbox -Identity $ObjectId -IsExchangeCloudManaged $CloudManaged -ErrorAction Stop
        return
    }
    $uri = "$GraphRoot/$script:SoaApiVersion/$($GraphSegment[$Category])/$ObjectId/onPremisesSyncBehavior"
    $null = Invoke-GraphRequestWithRetry -Method PATCH -Uri $uri -Body @{ isCloudManaged = $CloudManaged }
}

function New-ResultRow {
    param([string]$ObjectId, [string]$DisplayName, [string]$StateBefore, [string]$StateAfter, [string]$Status, [string]$ErrorMessage)
    return [pscustomobject][ordered]@{
        ObjectId        = $ObjectId
        DisplayName     = $DisplayName
        Category        = $Category
        StateBefore     = $StateBefore
        StateAfter      = $StateAfter
        Status          = $Status
        ErrorMessage    = $ErrorMessage
        Timestamp       = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        CurrentSoaState = $StateAfter
    }
}

function Invoke-VerifyLoop {
    # Polls submitted objects until they reach the expected state (eventual consistency).
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Pending, [Parameter(Mandatory)][string]$Expected)
    $results = [System.Collections.Generic.List[object]]::new()
    $open = [System.Collections.Generic.List[object]]@($Pending)
    $lastState = @{}
    for ($attempt = 1; $attempt -le $VerifyAttempts -and $open.Count -gt 0; $attempt++) {
        Write-Log "Verification $attempt/$VerifyAttempts for $($open.Count) object(s) in $VerifyDelaySeconds s..."
        Start-Sleep -Seconds $VerifyDelaySeconds
        foreach ($item in @($open)) {
            $current = Get-ObjectSoaState -ObjectId $item.ObjectId
            $lastState[$item.ObjectId] = $current
            if (Test-StateMatchesExpected -State $current.State -Expected $Expected) {
                $results.Add((New-ResultRow -ObjectId $item.ObjectId -DisplayName $item.DisplayName -StateBefore $item.StateBefore -StateAfter $current.State -Status 'Success' -ErrorMessage ''))
                $null = $open.Remove($item)
            }
        }
    }
    foreach ($item in $open) {
        $current = $lastState[$item.ObjectId]
        $message = "State not confirmed after $VerifyAttempts checks (expected $Expected, last state $($current.State)). $($current.Error)".Trim()
        $results.Add((New-ResultRow -ObjectId $item.ObjectId -DisplayName $item.DisplayName -StateBefore $item.StateBefore -StateAfter $current.State -Status 'Failed' -ErrorMessage $message))
    }
    return $results
}

function Show-ChangeWarning {
    param([Parameter(Mandatory)][string]$BackupPath)
    if ($Action -eq 'Convert') {
        switch ($Category) {
            'ExchangeAttributes' {
                Show-WarningBlock -Title 'Exchange attributes SOA conversion' -Lines @(
                    'After the last on-premises change (Set-RemoteMailbox) or mailbox move, wait one full sync cycle PLUS 24 hours for EVERY mailbox in the list (e.g. 6 h sync = 30 h in total).',
                    'After a mailbox move: first confirm the move completed and the on-premises recipient is a RemoteMailbox, let the post-move attributes sync to Entra ID, and only then start the 24-hour wait.',
                    'Converting too early can let stale cloud recipient-type values overwrite the correct remote-mailbox values in AD when Exchange attribute writeback runs.',
                    'This script cannot verify any of that. Do not continue if you are not sure.',
                    'After conversion, Exchange attributes stop updating from on-premises and are managed in Exchange Online. Identity attributes (name, UPN, department, ...) stay managed in AD.',
                    'Cloud Sync writes back only some attributes to AD (CustomAttribute1-15, ExtensionCustomAttribute1-5, mail, proxyAddresses, recipient type). Others, e.g. HiddenFromAddressListsEnabled or forwarding, stay cloud-only.'
                )
            }
            'User' {
                Show-WarningBlock -Title 'USER SOA conversion - the whole user object becomes cloud authoritative' -Lines @(
                    'Changes made in on-premises AD are no longer applied to these users (Connect Sync logs event 6956 instead).',
                    'User SOA conversion clears the IsExchangeCloudManaged state of the mailbox; the mailbox user is no longer dir-synced (IsDirSynced = False).',
                    'Sign-in: users federated through AD FS are NOT supported. Users must not depend on password-based on-premises authentication (LDAP bind, Kerberos with password).',
                    'Plan cloud authentication for these users (for example Windows Hello for Business with Cloud Kerberos Trust or FIDO2) and test sign-in with a pilot user first.',
                    'Do not convert users whose mailboxes are still on Exchange Server; all mailboxes must be in Exchange Online.',
                    'onPremises* attributes (SamAccountName, DistinguishedName, SID, ...) stay on the cloud object and must be maintained manually if on-premises access is still needed.'
                )
            }
            'Group' {
                Show-WarningBlock -Title 'Group SOA conversion' -Lines @(
                    'SOA is not inherited: nested groups stay on-premises managed. Convert each group separately, lowest level first.',
                    'Distribution groups and mail-enabled security groups are then managed only in Exchange Online; there is no writeback to AD.',
                    'extensionAttribute1-15 are not supported on converted groups.'
                )
            }
            'Contact' {
                Show-WarningBlock -Title 'Contact SOA conversion' -Lines @(
                    'Contact SOA requires Entra Connect Sync 2.5.79.0 or newer.',
                    'After conversion, on-premises changes to these contacts are no longer applied; manage them in Exchange Online.'
                )
            }
        }
    }
    else {
        $lines = @(
            'The next sync cycle OVERWRITES cloud values with the values from on-premises AD.',
            'This script does NOT restore attributes. Values worth keeping are in the backup CSV:',
            "  $BackupPath",
            'The rollback is complete only after the next (scheduled or forced) Connect Sync run.'
        )
        if ($Category -eq 'ExchangeAttributes') {
            $lines += 'Attributes that Cloud Sync does not write back (e.g. HiddenFromAddressListsEnabled, forwarding, moderation) revert to the AD values.'
            $lines += 'Mailbox permissions and junk e-mail settings are not in the backup CSV; export them separately if they were changed in the cloud.'
        }
        if ($Category -ne 'ExchangeAttributes') {
            $lines += 'Remove cloud references first: cloud-only members of converted groups, access package assignments.'
        }
        if ($Category -eq 'User') {
            $lines += 'Microsoft documents that user rollback requires the tenant feature blockCloudObjectTakeoverThroughHardMatchEnabled to be disabled, and re-enabled afterwards. This script does NOT change that tenant-wide setting.'
        }
        if ($Category -eq 'Group') {
            $lines += 'Groups with cloud-only members (CloudMemberCount > 0) are skipped by this script.'
        }
        Show-WarningBlock -Title "ROLLBACK of $Category SOA" -Lines $lines
    }
}

function Invoke-SoaChange {
    param([Parameter(Mandatory)][System.Management.Automation.PSCmdlet]$Cmdlet, [Parameter(Mandatory)][string]$RunFolder)
    $targetValue = $Action -eq 'Convert'
    $flagName = if ($Category -eq 'ExchangeAttributes') { 'IsExchangeCloudManaged' } else { 'isCloudManaged' }
    $expected = if ($targetValue) { 'CloudManaged' } else { 'OnPremManaged' }

    $inputRows = Read-InputCsv -Path $InputCsv -RequireCategory
    Write-Log "Input CSV: $($inputRows.Count) object(s)."

    Write-Log 'Taking a fresh inventory/backup of the objects before any change...'
    $inventory = Invoke-SoaInventory -ObjectIds @($inputRows.ObjectId)
    $backupPath = Export-SoaCsv -Data @($inventory.Rows) -Path (Join-Path $RunFolder "Backup_Inventory_$Category.csv")
    if ($Category -eq 'Group') {
        $null = Export-SoaCsv -Data @($inventory.Memberships) -Path (Join-Path $RunFolder 'Backup_GroupMemberships.csv')
    }
    if (-not $backupPath) { throw 'Backup could not be written. Nothing was changed.' }

    $results = [System.Collections.Generic.List[object]]::new()
    $toChange = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $inventory.Rows) {
        $skipReason = $null
        if ($Action -eq 'Convert') {
            if (-not $row.EligibleForConvert) { $skipReason = $row.Note }
        }
        elseif (Test-ProtectedRow -Row $row) {
            $skipReason = $row.Note
        }
        else {
            switch ($row.CurrentSoaState) {
                'CloudManaged' {
                    if ($Category -eq 'Group' -and [int]$row.CloudMemberCount -gt 0) {
                        $skipReason = "HasCloudMembers: $($row.CloudMemberCount) member(s) are cloud objects (cloud-only or SOA-converted); remove them before rollback."
                    }
                    elseif ($Category -eq 'Group' -and $row.GroupType -notin (Get-SelectedGroupKinds)) {
                        $skipReason = "GroupTypeNotSelected: group type is $($row.GroupType), -GroupType is $GroupType"
                    }
                }
                'OnPremManaged' { $skipReason = 'AlreadyOnPremManaged' }
                default         { $skipReason = $row.Note }
            }
        }
        if ($skipReason) {
            $results.Add((New-ResultRow -ObjectId $row.ObjectId -DisplayName $row.DisplayName -StateBefore $row.CurrentSoaState -StateAfter $row.CurrentSoaState -Status 'Skipped' -ErrorMessage $skipReason))
        }
        else {
            $toChange.Add($row)
        }
    }

    Show-ChangeWarning -BackupPath $backupPath
    Write-Host ''
    Write-Host "  Action          : $Action ($flagName -> $targetValue)" -ForegroundColor Cyan
    Write-Host "  Category        : $Category$(if ($Category -eq 'Group') { " (GroupType: $GroupType)" })" -ForegroundColor Cyan
    Write-Host "  Objects in CSV  : $($inventory.Rows.Count)" -ForegroundColor Cyan
    Write-Host "  To change       : $($toChange.Count)" -ForegroundColor Cyan
    Write-Host "  Skipped         : $($results.Count)" -ForegroundColor Cyan
    foreach ($group in ($results | Group-Object { ($_.ErrorMessage -split ':')[0] })) {
        Write-Host "      $($group.Count) x $($group.Name)" -ForegroundColor Cyan
    }
    Write-Host ''

    if ($toChange.Count -gt 0 -and -not $WhatIfPreference) {
        $confirmed = $true
        if ($Action -eq 'Convert' -and $Category -eq 'User') {
            $confirmed = Confirm-TypedText -Prompt 'I have read the USER SOA warning above and understand the impact on sign-in and AD synchronization.'
        }
        if ($confirmed) {
            $confirmed = Confirm-TypedText -Prompt "$Action $($toChange.Count) $Category object(s)?"
        }
        if (-not $confirmed) {
            Write-Log 'Cancelled by operator. Nothing was changed.' -Level Warn
            foreach ($row in $toChange) {
                $results.Add((New-ResultRow -ObjectId $row.ObjectId -DisplayName $row.DisplayName -StateBefore $row.CurrentSoaState -StateAfter $row.CurrentSoaState -Status 'Skipped' -ErrorMessage 'CancelledByOperator'))
            }
            $toChange.Clear()
        }
    }
    elseif ($WhatIfPreference) {
        Write-Log 'WhatIf: no confirmation requested and nothing will be changed.' -Level Warn
    }

    $submitted = [System.Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($row in $toChange) {
        $index++
        $label = "$($row.DisplayName) ($($row.ObjectId))"
        if (-not $Cmdlet.ShouldProcess($label, "Set $flagName = $targetValue")) {
            $results.Add((New-ResultRow -ObjectId $row.ObjectId -DisplayName $row.DisplayName -StateBefore $row.CurrentSoaState -StateAfter $row.CurrentSoaState -Status 'Skipped' -ErrorMessage $(if ($WhatIfPreference) { 'WhatIf' } else { 'DeclinedByOperator' })))
            continue
        }
        try {
            Set-ObjectSoa -ObjectId $row.ObjectId -CloudManaged $targetValue
            Write-Log "[$index/$($toChange.Count)] $label : $flagName set to $targetValue" -Level Success
            $submitted.Add([pscustomobject]@{ ObjectId = $row.ObjectId; DisplayName = $row.DisplayName; StateBefore = $row.CurrentSoaState })
        }
        catch {
            $message = Get-ErrorText -ErrorRecord $_
            Write-Log "[$index/$($toChange.Count)] $label : FAILED - $message" -Level Error
            $results.Add((New-ResultRow -ObjectId $row.ObjectId -DisplayName $row.DisplayName -StateBefore $row.CurrentSoaState -StateAfter $row.CurrentSoaState -Status 'Failed' -ErrorMessage $message))
        }
    }

    if ($submitted.Count -gt 0) {
        foreach ($result in (Invoke-VerifyLoop -Pending @($submitted) -Expected $expected)) { $results.Add($result) }
    }
    return [pscustomobject]@{ Results = $results; BackupPath = $backupPath }
}

function Invoke-SoaVerify {
    $inputRows = Read-InputCsv -Path $InputCsv
    Write-Log "Verifying $($inputRows.Count) object(s), expected state: $ExpectedState"
    $results = [System.Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($row in $inputRows) {
        $index++
        Write-Progress -Activity "Verify $Category" -Status "$index / $($inputRows.Count)" -PercentComplete ($index * 100 / $inputRows.Count)
        $before = if ($row.PSObject.Properties['CurrentSoaState']) { $row.CurrentSoaState } elseif ($row.PSObject.Properties['StateAfter']) { $row.StateAfter } else { '' }
        $current = Get-ObjectSoaState -ObjectId $row.ObjectId
        $status, $message = if ($current.Error) { 'Failed', $current.Error }
        elseif ($current.State -eq 'NotFound') { 'Skipped', 'NotFound' }
        elseif (Test-StateMatchesExpected -State $current.State -Expected $ExpectedState) { 'Success', '' }
        else { 'Failed', "Expected $ExpectedState, found $($current.State)" }
        $results.Add((New-ResultRow -ObjectId $row.ObjectId -DisplayName $current.DisplayName -StateBefore $before -StateAfter $current.State -Status $status -ErrorMessage $message))
    }
    Write-Progress -Activity "Verify $Category" -Completed
    return $results
}

#endregion

#region Main

if ($Action -ne 'Inventory' -and -not $InputCsv) {
    throw "-InputCsv is required for -Action $Action."
}
if ($PSBoundParameters.ContainsKey('GroupType') -and $Category -ne 'Group') {
    throw '-GroupType can only be used with -Category Group.'
}
if ($PSBoundParameters.ContainsKey('ExpectedState') -and $Action -ne 'Verify') {
    throw '-ExpectedState can only be used with -Action Verify.'
}

$needGraph = $Category -ne 'ExchangeAttributes'
$needExo = $Category -eq 'ExchangeAttributes' -or
    ($Action -ne 'Verify' -and ($Category -eq 'Contact' -or ($Category -eq 'Group' -and $GroupType -ne 'Security')))

$runFolder = New-RunFolder
$transcriptPath = Join-Path $runFolder 'Transcript.log'
Start-Transcript -LiteralPath $transcriptPath -WhatIf:$false -Confirm:$false | Out-Null

$exitCode = 0
$writtenFiles = [System.Collections.Generic.List[string]]::new()
$writtenFiles.Add($transcriptPath)
try {
    Write-Log "Invoke-SoaConversion: Action=$Action Category=$Category$(if ($Category -eq 'Group') { " GroupType=$GroupType" }) WhatIf=$WhatIfPreference"
    Write-Log "Output folder: $runFolder"

    Test-Prerequisites -NeedGraph $needGraph -NeedExo $needExo
    # Exchange Online first: loading it after Microsoft.Graph can cause MSAL assembly conflicts.
    $exoConnection = if ($needExo) { Connect-SoaExchange }
    $graphContext = if ($needGraph) { Connect-SoaGraph -Scopes $GraphScopes[$Category] }
    if ($Category -eq 'User' -and $Action -ne 'Verify') { Initialize-ProtectedObjects }
    if ($exoConnection -and $graphContext -and "$($exoConnection.TenantID)" -ne "$($graphContext.TenantId)") {
        throw "Exchange Online tenant ($($exoConnection.TenantID)) and Microsoft Graph tenant ($($graphContext.TenantId)) differ."
    }

    switch ($Action) {
        'Inventory' {
            $ids = if ($InputCsv) { @((Read-InputCsv -Path $InputCsv).ObjectId) } else { $null }
            $inventory = Invoke-SoaInventory -ObjectIds $ids
            $path = Export-SoaCsv -Data @($inventory.Rows) -Path (Join-Path $runFolder "Inventory_$Category.csv")
            if ($path) { $writtenFiles.Add($path) }
            if ($Category -eq 'Group') {
                $path = Export-SoaCsv -Data @($inventory.Memberships) -Path (Join-Path $runFolder 'GroupMemberships.csv')
                if ($path) { $writtenFiles.Add($path) }
            }
            $eligible = @($inventory.Rows | Where-Object EligibleForConvert).Count
            Write-Host ''
            Write-Log "Inventory: $($inventory.Rows.Count) object(s), eligible for conversion: $eligible, not eligible: $($inventory.Rows.Count - $eligible)" -Level Success
            foreach ($group in ($inventory.Rows | Where-Object { -not $_.EligibleForConvert } | Group-Object { ($_.Note -split ':')[0] })) {
                Write-Log "  $($group.Count) x $($group.Name)"
            }
        }
        { $_ -in 'Convert', 'Rollback' } {
            $change = Invoke-SoaChange -Cmdlet $PSCmdlet -RunFolder $runFolder
            if ($change.BackupPath) { $writtenFiles.Add($change.BackupPath) }
            $groupBackup = Join-Path $runFolder 'Backup_GroupMemberships.csv'
            if (Test-Path -LiteralPath $groupBackup) { $writtenFiles.Add($groupBackup) }
            $results = @($change.Results)
        }
        'Verify' {
            $results = @(Invoke-SoaVerify)
        }
    }

    if ($Action -ne 'Inventory') {
        $path = Export-SoaCsv -Data $results -Path (Join-Path $runFolder "Results_${Action}_$Category.csv")
        if ($path) { $writtenFiles.Add($path) }
        $success = @($results | Where-Object Status -eq 'Success').Count
        $failed = @($results | Where-Object Status -eq 'Failed').Count
        $skipped = @($results | Where-Object Status -eq 'Skipped').Count
        Write-Host ''
        Write-Log "Summary: Success $success | Failed $failed | Skipped $skipped" -Level $(if ($failed) { 'Error' } else { 'Success' })
        if ($failed) { $exitCode = 1 }
    }
    if ($needGraph) { Write-Log "Graph SOA endpoint version used: $script:SoaApiVersion" }
}
catch {
    Write-Log "Stopped: $(Get-ErrorText -ErrorRecord $_)" -Level Error
    $exitCode = 2
}
finally {
    Write-Log 'Files:'
    foreach ($file in $writtenFiles) { Write-Log "  $file" }
    Stop-Transcript -WhatIf:$false -Confirm:$false | Out-Null
}
exit $exitCode

#endregion
