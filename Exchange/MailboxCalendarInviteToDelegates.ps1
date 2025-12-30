<#
.SYNOPSIS
    Fixes calendar invite delegation so the mailbox owner receives invites, not delegates
.DESCRIPTION
    When a mailbox has delegates receiving calendar invites instead of the owner,
    this script corrects the calendar processing settings
.NOTES
    Requires Exchange Online PowerShell module and appropriate permissions
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$MailboxIdentity,
    
    [Parameter(Mandatory = $false)]
    [string]$AdminUPN,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " FIX CALENDAR INVITE DELEGATION" -ForegroundColor Cyan
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

# Get mailbox identity if not provided
if (-not $MailboxIdentity) {
    $MailboxIdentity = Read-Host "Enter mailbox identity (email address or display name)"
}

# Resolve mailbox
Write-Host "Resolving mailbox: $MailboxIdentity" -ForegroundColor Cyan
try {
    $Mailbox = Get-EXOMailbox -Identity $MailboxIdentity -ErrorAction Stop
    Write-Host "Found: $($Mailbox.DisplayName) ($($Mailbox.PrimarySmtpAddress))" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Could not find mailbox '$MailboxIdentity'" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

Write-Host ""
Write-Host "Current Calendar Processing Settings:" -ForegroundColor Cyan

# Get current calendar processing settings
try {
    $CalSettings = Get-CalendarProcessing -Identity $Mailbox.PrimarySmtpAddress
    
    Write-Host "  AutomateProcessing: $($CalSettings.AutomateProcessing)" -ForegroundColor White
    Write-Host "  ForwardRequestsToDelegates: $($CalSettings.ForwardRequestsToDelegates)" -ForegroundColor $(if ($CalSettings.ForwardRequestsToDelegates) { 'Red' } else { 'Green' })
    Write-Host "  ProcessExternalMeetingMessages: $($CalSettings.ProcessExternalMeetingMessages)" -ForegroundColor White
    
    if ($CalSettings.ResourceDelegates) {
        Write-Host "  ResourceDelegates:" -ForegroundColor White
        $CalSettings.ResourceDelegates | ForEach-Object { Write-Host "    - $_" -ForegroundColor Yellow }
    } else {
        Write-Host "  ResourceDelegates: None" -ForegroundColor White
    }
    
    if ($CalSettings.ForwardRequestsToDelegates) {
        Write-Host ""
        Write-Host "ISSUE FOUND: ForwardRequestsToDelegates is set to TRUE" -ForegroundColor Red
        Write-Host "This causes delegates to receive invites instead of the mailbox owner." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "ERROR: Could not retrieve calendar settings: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

Write-Host ""

if ($WhatIf) {
    Write-Host "WHAT-IF MODE: No actual changes will be made" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Would apply these settings:" -ForegroundColor Cyan
    Write-Host "  ForwardRequestsToDelegates: FALSE" -ForegroundColor Green
    Write-Host "  AutomateProcessing: AutoUpdate (ensures owner gets invites)" -ForegroundColor Green
    Write-Host ""
    Write-Host "This will ensure $($Mailbox.DisplayName) receives calendar invites directly." -ForegroundColor White
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 0
}

# Confirmation
Write-Host "This will configure the mailbox so:" -ForegroundColor Cyan
Write-Host "  - Calendar invites go to $($Mailbox.DisplayName) (owner)" -ForegroundColor White
Write-Host "  - Delegates will NOT receive invites automatically" -ForegroundColor White
Write-Host "  - Delegates can still view/edit calendar if they have permissions" -ForegroundColor White
Write-Host ""
$Confirmation = Read-Host "Type 'YES' to proceed, or anything else to cancel"
if ($Confirmation -ne "YES") {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 0
}

# Apply fix
Write-Host ""
Write-Host "Applying calendar processing fix..." -ForegroundColor Cyan

try {
    Set-CalendarProcessing -Identity $Mailbox.PrimarySmtpAddress `
        -ForwardRequestsToDelegates $false `
        -AutomateProcessing AutoUpdate `
        -ErrorAction Stop
    
    Write-Host "SUCCESS: Calendar settings updated" -ForegroundColor Green
    Write-Host ""
    
    # Verify the change
    Write-Host "Verifying new settings..." -ForegroundColor Cyan
    $NewSettings = Get-CalendarProcessing -Identity $Mailbox.PrimarySmtpAddress
    Write-Host "  ForwardRequestsToDelegates: $($NewSettings.ForwardRequestsToDelegates)" -ForegroundColor Green
    Write-Host "  AutomateProcessing: $($NewSettings.AutomateProcessing)" -ForegroundColor Green
    Write-Host ""
    Write-Host "$($Mailbox.DisplayName) will now receive calendar invites directly." -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to update calendar settings: $($_.Exception.Message)" -ForegroundColor Red
}

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host ""
Write-Host "Disconnected from Exchange Online." -ForegroundColor DarkGray
Write-Host "Operation complete." -ForegroundColor Green

<#
.USAGE EXAMPLES

Check current settings (no changes):
.\MailboxCalendarInviteToDelegates.ps1 -MailboxIdentity "user@domain.com" -WhatIf

Fix calendar invite delegation:
.\MailboxCalendarInviteToDelegates.ps1 -MailboxIdentity "user@domain.com"

Interactive mode (prompts for mailbox):
.\MailboxCalendarInviteToDelegates.ps1
#>