<#
.SYNOPSIS
    Adds additional email addresses to mailboxes with disabled Email Address Policy.

.DESCRIPTION
    This script finds all on-premises Exchange mailboxes that have Email Address Policy
    disabled (EmailAddressPolicyEnabled = $false) and adds an additional email address
    using the specified domain suffix.

.PARAMETER DomainSuffix
    The domain suffix to use for new email addresses (e.g., @company.mail.onmicrosoft.com).
    This parameter is mandatory.

.PARAMETER WhatIf
    Shows what would happen if the script runs without actually making changes.

.PARAMETER Confirm
    Prompts for confirmation before making changes.

.EXAMPLE
    .\Add-SecondaryEmailAddress.ps1 -DomainSuffix "@company.mail.onmicrosoft.com"

    Adds email addresses in format alias@company.mail.onmicrosoft.com to all qualifying mailboxes.

.EXAMPLE
    .\Add-SecondaryEmailAddress.ps1 -DomainSuffix "@custom.onmicrosoft.com" -WhatIf

    Shows what changes would be made without actually making them.

.NOTES
    Author: PowerShell Script
    Date: 2025-12-04
    Requires: Exchange Management Shell
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter domain suffix (e.g., @company.mail.onmicrosoft.com)")]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern('^@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
    [string]$DomainSuffix,

    [Parameter()]
    [switch]$SkipExisting
)

# Initialize counters
$processedCount = 0
$skippedCount = 0
$errorCount = 0
$results = @()

Write-Host "Starting email address addition process..." -ForegroundColor Cyan
Write-Host "Domain suffix: $DomainSuffix" -ForegroundColor Cyan
Write-Host ("-" * 80) -ForegroundColor Gray

try {
    # Check if Exchange Management Shell is loaded
    if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        throw "Exchange Management Shell is not loaded. Please run this script from Exchange Management Shell or load the Exchange module."
    }

    # Get all mailboxes with Email Address Policy disabled
    Write-Host "Retrieving mailboxes with Email Address Policy disabled..." -ForegroundColor Yellow

    $mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {
        $_.EmailAddressPolicyEnabled -eq $false
    }

    $totalMailboxes = ($mailboxes | Measure-Object).Count
    Write-Host "Found $totalMailboxes mailbox(es) with Email Address Policy disabled." -ForegroundColor Green
    Write-Host ("-" * 80) -ForegroundColor Gray

    if ($totalMailboxes -eq 0) {
        Write-Host "No mailboxes found. Exiting." -ForegroundColor Yellow
        return
    }

    # Process each mailbox
    foreach ($mailbox in $mailboxes) {
        $currentCount = $processedCount + $skippedCount + $errorCount + 1
        Write-Progress -Activity "Processing mailboxes" `
                       -Status "Processing $currentCount of $totalMailboxes" `
                       -PercentComplete (($currentCount / $totalMailboxes) * 100)

        $alias = $mailbox.Alias
        $displayName = $mailbox.DisplayName
        $primaryEmail = $mailbox.PrimarySmtpAddress
        $newEmailAddress = "smtp:$alias$DomainSuffix"

        Write-Host "[$currentCount/$totalMailboxes] Processing: $displayName ($primaryEmail)" -ForegroundColor White

        try {
            # Check if the email address already exists
            $existingAddresses = $mailbox.EmailAddresses | ForEach-Object { $_.ToString() }
            $addressExists = $existingAddresses | Where-Object {
                $_ -like "*$alias$DomainSuffix"
            }

            if ($addressExists) {
                Write-Host "  [SKIP] Email address already exists: $newEmailAddress" -ForegroundColor Yellow
                $skippedCount++

                $results += [PSCustomObject]@{
                    DisplayName      = $displayName
                    Alias           = $alias
                    PrimaryEmail    = $primaryEmail
                    NewEmailAddress = $newEmailAddress
                    Status          = "Skipped - Already exists"
                    Timestamp       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }

                continue
            }

            # Add the new email address
            # Note: We must read all addresses and write them back because Exchange in production
            # may run in ConstrainedLanguage mode which doesn't support @{Add=...} syntax
            if ($PSCmdlet.ShouldProcess($displayName, "Add email address $newEmailAddress")) {
                try {
                    # Get current addresses (fresh read to avoid stale data)
                    $currentMailbox = Get-Mailbox -Identity $mailbox.Identity -ErrorAction Stop
                    $allAddresses = [System.Collections.ArrayList]::new()

                    # Copy existing addresses to new collection
                    foreach ($addr in $currentMailbox.EmailAddresses) {
                        [void]$allAddresses.Add($addr.ToString())
                    }

                    # Add the new address
                    [void]$allAddresses.Add($newEmailAddress)

                    # Write all addresses back (atomic operation)
                    Set-Mailbox -Identity $currentMailbox.Identity -EmailAddresses $allAddresses -ErrorAction Stop

                    Write-Host "  [SUCCESS] Added: $newEmailAddress" -ForegroundColor Green
                    $processedCount++

                    $results += [PSCustomObject]@{
                        DisplayName      = $displayName
                        Alias           = $alias
                        PrimaryEmail    = $primaryEmail
                        NewEmailAddress = $newEmailAddress
                        Status          = "Success"
                        Timestamp       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                catch {
                    throw  # Re-throw to be caught by outer catch block
                }
            }
        }
        catch {
            Write-Host "  [ERROR] Failed to add email address: $($_.Exception.Message)" -ForegroundColor Red
            $errorCount++

            $results += [PSCustomObject]@{
                DisplayName      = $displayName
                Alias           = $alias
                PrimaryEmail    = $primaryEmail
                NewEmailAddress = $newEmailAddress
                Status          = "Error: $($_.Exception.Message)"
                Timestamp       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }

        Write-Host ""
    }

    Write-Progress -Activity "Processing mailboxes" -Completed

}
catch {
    Write-Host "CRITICAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
}
finally {
    # Summary
    Write-Host ("-" * 80) -ForegroundColor Gray
    Write-Host "SUMMARY" -ForegroundColor Cyan
    Write-Host ("-" * 80) -ForegroundColor Gray
    Write-Host "Total mailboxes found:     $totalMailboxes" -ForegroundColor White
    Write-Host "Successfully processed:    $processedCount" -ForegroundColor Green
    Write-Host "Skipped (already exists):  $skippedCount" -ForegroundColor Yellow
    Write-Host "Errors:                    $errorCount" -ForegroundColor Red
    Write-Host ("-" * 80) -ForegroundColor Gray

    # Export results to CSV
    if ($results.Count -gt 0) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $reportPath = Join-Path $PSScriptRoot "EmailAddressReport_$timestamp.csv"

        try {
            $results | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
            Write-Host "Report exported to: $reportPath" -ForegroundColor Cyan
        }
        catch {
            Write-Host "Failed to export report: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host "Script completed." -ForegroundColor Cyan
}
