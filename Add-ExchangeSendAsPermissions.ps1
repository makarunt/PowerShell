<#
.SYNOPSIS
    Adds SendAs permissions to Exchange Online distribution groups from a CSV file.

.DESCRIPTION
    This script reads a CSV file containing distribution group identities (PrimarySmtpAddress)
    and user UPNs, then adds SendAs permissions accordingly. All operations are logged,
    with successes and failures written to separate log files.

.PARAMETER CsvPath
    Path to the input CSV file. The CSV must contain 'PrimarySmtpAddress' and 'UPN' columns.

.PARAMETER LogDirectory
    Directory where log files will be created. Defaults to the script's directory.

.EXAMPLE
    .\Add-ExchangeSendAsPermissions.ps1 -CsvPath "C:\Data\sendas-permissions.csv"

.EXAMPLE
    .\Add-ExchangeSendAsPermissions.ps1 -CsvPath ".\permissions.csv" -LogDirectory "C:\Logs"

.NOTES
    Requires Exchange Online PowerShell module (ExchangeOnlineManagement).
    User must have appropriate Exchange Online admin permissions.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter()]
    [string]$LogDirectory = $PSScriptRoot
)

#region Functions

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory)]
        [string]$LogFile,

        [Parameter()]
        [ValidateSet('INFO', 'SUCCESS', 'ERROR', 'WARNING')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"

    Add-Content -Path $LogFile -Value $logEntry -Encoding UTF8

    switch ($Level) {
        'ERROR'   { Write-Host $logEntry -ForegroundColor Red }
        'WARNING' { Write-Host $logEntry -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
        default   { Write-Host $logEntry }
    }
}

function Test-ExchangeOnlineConnection {
    [CmdletBinding()]
    param()

    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

#endregion Functions

#region Main Script

# Create timestamp for log files
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$successLogFile = Join-Path -Path $LogDirectory -ChildPath "SendAs_Success_$timestamp.log"
$errorLogFile = Join-Path -Path $LogDirectory -ChildPath "SendAs_Errors_$timestamp.log"

# Ensure log directory exists
if (-not (Test-Path -Path $LogDirectory -PathType Container)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

# Initialize log files
"SendAs Permission Addition - Success Log - Started: $(Get-Date)" | Out-File -FilePath $successLogFile -Encoding UTF8
"SendAs Permission Addition - Error Log - Started: $(Get-Date)" | Out-File -FilePath $errorLogFile -Encoding UTF8

Write-Log -Message "Script started" -LogFile $successLogFile -Level INFO
Write-Log -Message "CSV Path: $CsvPath" -LogFile $successLogFile -Level INFO
Write-Log -Message "Success Log: $successLogFile" -LogFile $successLogFile -Level INFO
Write-Log -Message "Error Log: $errorLogFile" -LogFile $successLogFile -Level INFO

# Check Exchange Online connection
if (-not (Test-ExchangeOnlineConnection)) {
    Write-Log -Message "Not connected to Exchange Online. Attempting to connect..." -LogFile $successLogFile -Level WARNING

    try {
        Connect-ExchangeOnline -ErrorAction Stop
        Write-Log -Message "Successfully connected to Exchange Online" -LogFile $successLogFile -Level SUCCESS
    }
    catch {
        $errorMessage = "Failed to connect to Exchange Online: $($_.Exception.Message)"
        Write-Log -Message $errorMessage -LogFile $errorLogFile -Level ERROR
        throw $errorMessage
    }
}
else {
    Write-Log -Message "Already connected to Exchange Online" -LogFile $successLogFile -Level INFO
}

# Import CSV
try {
    $permissions = Import-Csv -Path $CsvPath -Delimiter ';' -ErrorAction Stop
    Write-Log -Message "Successfully imported CSV with $($permissions.Count) entries" -LogFile $successLogFile -Level SUCCESS
}
catch {
    $errorMessage = "Failed to import CSV file: $($_.Exception.Message)"
    Write-Log -Message $errorMessage -LogFile $errorLogFile -Level ERROR
    throw $errorMessage
}

# Validate CSV columns
$requiredColumns = @('PrimarySmtpAddress', 'UPN')
$csvColumns = $permissions[0].PSObject.Properties.Name

foreach ($column in $requiredColumns) {
    if ($column -notin $csvColumns) {
        $errorMessage = "CSV file is missing required column: $column"
        Write-Log -Message $errorMessage -LogFile $errorLogFile -Level ERROR
        throw $errorMessage
    }
}

Write-Log -Message "CSV validation passed" -LogFile $successLogFile -Level INFO

# Process each permission entry
$successCount = 0
$errorCount = 0
$totalCount = $permissions.Count

foreach ($entry in $permissions) {
    $distributionGroup = $entry.PrimarySmtpAddress
    $userUpn = $entry.UPN

    if ([string]::IsNullOrWhiteSpace($distributionGroup) -or [string]::IsNullOrWhiteSpace($userUpn)) {
        $errorMessage = "Skipping entry with empty values - DG: '$distributionGroup', UPN: '$userUpn'"
        Write-Log -Message $errorMessage -LogFile $errorLogFile -Level WARNING
        $errorCount++
        continue
    }

    try {
        # Add SendAs permission using Add-RecipientPermission
        Add-RecipientPermission -Identity $distributionGroup -Trustee $userUpn -AccessRights SendAs -Confirm:$false -ErrorAction Stop

        $successMessage = "Successfully added SendAs permission - DG: '$distributionGroup', Trustee: '$userUpn'"
        Write-Log -Message $successMessage -LogFile $successLogFile -Level SUCCESS
        $successCount++
    }
    catch {
        $errorMessage = "Failed to add SendAs permission - DG: '$distributionGroup', Trustee: '$userUpn' - Error: $($_.Exception.Message)"
        Write-Log -Message $errorMessage -LogFile $errorLogFile -Level ERROR
        $errorCount++
    }
}

# Summary
$summaryMessage = @"
========================================
Script completed
Total entries processed: $totalCount
Successful: $successCount
Failed: $errorCount
========================================
"@

Write-Log -Message $summaryMessage -LogFile $successLogFile -Level INFO

if ($errorCount -gt 0) {
    Write-Log -Message "There were $errorCount errors. Check error log: $errorLogFile" -LogFile $successLogFile -Level WARNING
}

Write-Host "`nScript completed. Check logs for details:"
Write-Host "  Success log: $successLogFile" -ForegroundColor Green
Write-Host "  Error log: $errorLogFile" -ForegroundColor Yellow

#endregion Main Script
