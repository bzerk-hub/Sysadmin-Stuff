<#
.SYNOPSIS
    Adds full access permissions between a group of users mailboxes
.DESCRIPTION
    Each user in the list gets full access to all other users mailboxes in the group
.NOTES
    Requires Exchange Online PowerShell module and appropriate permissions
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$AdminUPN,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

# Define the user group
$UserGroup = @(
    "Tania Kimball",
    "Rominy Morgan", 
    "Vicky Avery",
    "Julie Kerry",
    "Karen Ryan",
    "Husna Farooq",
    "Nisa Nazemi-Sandry",
    "Siobhan Stretton"
)

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " MAILBOX FULL ACCESS PERMISSION SETUP" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Install and import Exchange Online module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing Exchange Online Management module..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser -AllowClobber
}
Import-Module ExchangeOnlineManagement -Force

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
if ($AdminUPN) {
    Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false
} else {
    Connect-ExchangeOnline -ShowBanner:$false
}
Write-Host "Connected." -ForegroundColor Green
Write-Host ""

# ============================================
# RESOLVE USER IDENTITIES
# ============================================
Write-Host "[1/4] Resolving User Identities..." -ForegroundColor Cyan

$ResolvedUsers = @()
$FailedUsers = @()

foreach ($DisplayName in $UserGroup) {
    Write-Host "  Resolving: $DisplayName" -ForegroundColor White
    
    try {
        # Try to find by display name first
        $User = Get-EXOMailbox -Filter "DisplayName -eq '$DisplayName'" -ErrorAction Stop
        
        if ($User.Count -gt 1) {
            Write-Host "    Warning: Multiple matches found for '$DisplayName'. Using first match." -ForegroundColor Yellow
            $User = $User[0]
        }
        
        if ($User) {
            $ResolvedUsers += [PSCustomObject]@{
                DisplayName = $DisplayName
                PrimarySmtpAddress = $User.PrimarySmtpAddress
                UserPrincipalName = $User.UserPrincipalName
                Identity = $User.Identity
            }
            Write-Host "    + Found: $($User.PrimarySmtpAddress)" -ForegroundColor Green
        } else {
            throw "No mailbox found"
        }
    }
    catch {
        Write-Host "    X Failed: $($_.Exception.Message)" -ForegroundColor Red
        $FailedUsers += $DisplayName
    }
}

Write-Host ""
Write-Host "Resolution Summary:" -ForegroundColor Cyan
Write-Host "  Successfully resolved: $($ResolvedUsers.Count)" -ForegroundColor Green
Write-Host "  Failed to resolve: $($FailedUsers.Count)" -ForegroundColor $(if ($FailedUsers.Count -gt 0) { 'Red' } else { 'Green' })

if ($FailedUsers.Count -gt 0) {
    Write-Host "  Failed users:" -ForegroundColor Red
    foreach ($Failed in $FailedUsers) {
        Write-Host "    - $Failed" -ForegroundColor Red
    }
}

if ($ResolvedUsers.Count -lt 2) {
    Write-Host ""
    Write-Host "ERROR: Need at least 2 resolved users to proceed." -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

Write-Host ""

# ============================================
# SHOW PERMISSIONS MATRIX
# ============================================
Write-Host "[2/4] Permission Matrix to be Applied..." -ForegroundColor Cyan

$PermissionMatrix = @()

foreach ($MailboxOwner in $ResolvedUsers) {
    foreach ($Trustee in $ResolvedUsers) {
        if ($MailboxOwner.Identity -ne $Trustee.Identity) {
            $PermissionMatrix += [PSCustomObject]@{
                Mailbox = $MailboxOwner.DisplayName
                MailboxEmail = $MailboxOwner.PrimarySmtpAddress
                Trustee = $Trustee.DisplayName
                TrusteeEmail = $Trustee.PrimarySmtpAddress
                Permission = "FullAccess"
                Status = "Pending"
            }
        }
    }
}

Write-Host "Total permissions to be granted: $($PermissionMatrix.Count)" -ForegroundColor White
Write-Host ""
Write-Host "Preview (first 10):" -ForegroundColor Yellow
$PermissionMatrix | Select-Object -First 10 | Format-Table Mailbox, Trustee, Permission -AutoSize

if ($WhatIf) {
    Write-Host ""
    Write-Host "WHAT-IF MODE: No actual changes will be made" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Full permission matrix that WOULD be applied:" -ForegroundColor Cyan
    $PermissionMatrix | Format-Table -AutoSize
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 0
}

# ============================================
# CONFIRMATION
# ============================================
Write-Host ""
Write-Host "WARNING: This will grant FullAccess permissions as shown above!" -ForegroundColor Red
Write-Host ""
$Confirmation = Read-Host "Type 'YES' to proceed, or anything else to cancel"
if ($Confirmation -ne "YES") {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 0
}

# ============================================
# APPLY PERMISSIONS
# ============================================
Write-Host ""
Write-Host "[3/4] Applying Full Access Permissions..." -ForegroundColor Cyan

$SuccessCount = 0
$SkippedCount = 0
$FailureCount = 0

$Counter = 0
foreach ($Permission in $PermissionMatrix) {
    $Counter++
    $Progress = [math]::Round(($Counter / $PermissionMatrix.Count) * 100, 1)
    
    Write-Host ""
    Write-Host "[$Counter/$($PermissionMatrix.Count)] ($Progress%) Processing: $($Permission.Trustee) -> $($Permission.Mailbox)" -ForegroundColor Yellow
    
    try {
        # Check if permission already exists
        $ExistingPermission = Get-EXOMailboxPermission -Identity $Permission.MailboxEmail -User $Permission.TrusteeEmail -ErrorAction SilentlyContinue
        
        if ($ExistingPermission -and $ExistingPermission.AccessRights -contains "FullAccess") {
            Write-Host "  - Permission already exists, skipping" -ForegroundColor DarkGray
            $Permission.Status = "Already Exists"
            $SkippedCount++
        } else {
            # Add the permission
            Add-MailboxPermission -Identity $Permission.MailboxEmail -User $Permission.TrusteeEmail -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -ErrorAction Stop
            Write-Host "  + Permission granted successfully" -ForegroundColor Green
            $Permission.Status = "Success"
            $SuccessCount++
        }
    }
    catch {
        Write-Host "  X Failed: $($_.Exception.Message)" -ForegroundColor Red
        $Permission.Status = "Failed: $($_.Exception.Message)"
        $FailureCount++
    }
}

# ============================================
# VERIFICATION & SUMMARY
# ============================================
Write-Host ""
Write-Host "[4/4] Verification & Summary..." -ForegroundColor Cyan

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " OPERATION SUMMARY" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

Write-Host "Permissions granted: $SuccessCount" -ForegroundColor Green
Write-Host "Already existed: $SkippedCount" -ForegroundColor Yellow
Write-Host "Failed: $FailureCount" -ForegroundColor $(if ($FailureCount -gt 0) { 'Red' } else { 'Green' })
Write-Host "Total processed: $($PermissionMatrix.Count)" -ForegroundColor White

if ($FailureCount -gt 0) {
    Write-Host ""
    Write-Host "FAILED PERMISSIONS:" -ForegroundColor Red
    $FailedPermissions = $PermissionMatrix | Where-Object { $_.Status -like "Failed:*" }
    $FailedPermissions | Format-Table Mailbox, Trustee, Status -AutoSize
}

# Export results
$ExportPath = "C:\Temp\MailboxPermissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
if (-not (Test-Path "C:\Temp")) {
    New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null
}
$PermissionMatrix | Export-Csv -Path $ExportPath -NoTypeInformation
Write-Host ""
Write-Host "Detailed results exported to: $ExportPath" -ForegroundColor Green

# Sample verification
Write-Host ""
Write-Host "Sample verification (checking first user's permissions):" -ForegroundColor Cyan
if ($ResolvedUsers.Count -gt 0) {
    $SampleUser = $ResolvedUsers[0]
    Write-Host "Checking permissions on: $($SampleUser.DisplayName)" -ForegroundColor White
    
    try {
        $Permissions = Get-EXOMailboxPermission -Identity $SampleUser.PrimarySmtpAddress | Where-Object { $_.User -ne "NT AUTHORITY\SELF" -and $_.AccessRights -contains "FullAccess" }
        Write-Host "Full Access users found: $($Permissions.Count)" -ForegroundColor White
        foreach ($Perm in $Permissions) {
            Write-Host "  - $($Perm.User)" -ForegroundColor DarkGray
        }
    }
    catch {
        Write-Host "Could not verify sample permissions: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host ""
Write-Host "Disconnected from Exchange Online." -ForegroundColor DarkGray
Write-Host "Operation complete." -ForegroundColor Green

<#
.USAGE EXAMPLES

Test run (no changes):
.\Permission Add Delegates.ps1 -WhatIf

Apply permissions:
.\Permission Add Delegates.ps1

With specific admin account:
.\Permission Add Delegates.ps1 -AdminUPN "admin@domain.com"
#>