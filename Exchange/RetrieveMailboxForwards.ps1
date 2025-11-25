param(
    [Parameter(Mandatory = $false, HelpMessage = "Your admin account email (optional - will prompt if not provided)")]
    [string]$AdminUPN
)

$ErrorActionPreference = 'Stop'

try {
    # Import Exchange Online module
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser -AllowClobber
    }
    Import-Module ExchangeOnlineManagement -Force
    
    # Connect to Exchange Online
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host " Mailbox Forward Settings Report" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    if ($AdminUPN) {
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false
    } else {
        Connect-ExchangeOnline -ShowBanner:$false
    }
    Write-Host "Connected." -ForegroundColor Green
    Write-Host ""
    
    # Prompt user for scope
    Write-Host "Select scope:" -ForegroundColor Cyan
    Write-Host "  1) Check a specific mailbox" -ForegroundColor White
    Write-Host "  2) Check ALL mailboxes in organization" -ForegroundColor White
    Write-Host ""
    $Choice = Read-Host "Enter choice (1 or 2)"
    
    if ($Choice -eq "1") {
        Write-Host ""
        $MailboxInput = Read-Host "Enter mailbox email address"
        Write-Host ""
        Write-Host "Checking mailbox: $MailboxInput..." -ForegroundColor Cyan
        $Mailboxes = @(Get-EXOMailbox -Identity $MailboxInput -ErrorAction Stop)
        Write-Host "✓ Mailbox found: $($Mailboxes[0].DisplayName)" -ForegroundColor Green
    }
    elseif ($Choice -eq "2") {
        Write-Host ""
        Write-Host "Retrieving ALL mailboxes (this may take a few minutes)..." -ForegroundColor Yellow
        $Mailboxes = Get-EXOMailbox -ResultSize Unlimited
        Write-Host "✓ Found $($Mailboxes.Count) mailboxes to check." -ForegroundColor Green
    }
    else {
        throw "Invalid choice. Please enter 1 or 2."
    }
    Write-Host ""
    
    # Check forwarding settings
    $Results = @()
    $Counter = 0
    $Total = $Mailboxes.Count
    
    foreach ($Mbx in $Mailboxes) {
        $Counter++
        Write-Progress -Activity "Checking mailbox forwarding settings" -Status "Processing $Counter of $Total" -PercentComplete (($Counter / $Total) * 100)
        
        # Get full mailbox details including forwarding settings
        $MailboxDetails = Get-EXOMailbox -Identity $Mbx.PrimarySmtpAddress -Properties ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward
        
        # Get inbox rules that forward
        $InboxRules = Get-InboxRule -Mailbox $Mbx.PrimarySmtpAddress -ErrorAction SilentlyContinue | 
            Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo }
        
        # Check for forwarding address (internal)
        if ($MailboxDetails.ForwardingAddress) {
            $ForwardTo = (Get-Recipient -Identity $MailboxDetails.ForwardingAddress).PrimarySmtpAddress
            $Results += [PSCustomObject]@{
                Mailbox              = $MailboxDetails.DisplayName
                PrimarySmtpAddress   = $MailboxDetails.PrimarySmtpAddress
                ForwardType          = "Mailbox Forwarding (Internal)"
                ForwardTo            = $ForwardTo
                DeliverAndForward    = $MailboxDetails.DeliverToMailboxAndForward
                RuleName             = "N/A"
            }
        }
        
        # Check for forwarding SMTP address (external)
        if ($MailboxDetails.ForwardingSmtpAddress) {
            $ForwardTo = $MailboxDetails.ForwardingSmtpAddress -replace "smtp:", ""
            $Results += [PSCustomObject]@{
                Mailbox              = $MailboxDetails.DisplayName
                PrimarySmtpAddress   = $MailboxDetails.PrimarySmtpAddress
                ForwardType          = "Mailbox Forwarding (External)"
                ForwardTo            = $ForwardTo
                DeliverAndForward    = $MailboxDetails.DeliverToMailboxAndForward
                RuleName             = "N/A"
            }
        }
        
        # Check inbox rules
        foreach ($Rule in $InboxRules) {
            $ForwardAddresses = @()
            
            if ($Rule.ForwardTo) {
                $ForwardAddresses += $Rule.ForwardTo | ForEach-Object { $_.Split(']')[-1].Trim() }
            }
            if ($Rule.ForwardAsAttachmentTo) {
                $ForwardAddresses += $Rule.ForwardAsAttachmentTo | ForEach-Object { $_.Split(']')[-1].Trim() }
            }
            if ($Rule.RedirectTo) {
                $ForwardAddresses += $Rule.RedirectTo | ForEach-Object { $_.Split(']')[-1].Trim() }
            }
            
            foreach ($Address in $ForwardAddresses) {
                $Results += [PSCustomObject]@{
                    Mailbox              = $MailboxDetails.DisplayName
                    PrimarySmtpAddress   = $MailboxDetails.PrimarySmtpAddress
                    ForwardType          = "Inbox Rule"
                    ForwardTo            = $Address
                    DeliverAndForward    = if ($Rule.RedirectTo) { $false } else { $true }
                    RuleName             = $Rule.Name
                }
            }
        }
    }
    
    Write-Progress -Activity "Checking mailbox forwarding settings" -Completed
    
    # Display results
    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host " RESULTS" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    
    if ($Results.Count -eq 0) {
        Write-Host "✓ No forwarding configured on any mailboxes." -ForegroundColor Green
    } else {
        Write-Host "⚠ Found $($Results.Count) forwarding configuration(s):" -ForegroundColor Yellow
        Write-Host ""
        $Results | Format-Table -AutoSize
        
        # Ask if user wants to export
        Write-Host ""
        $ExportChoice = Read-Host "Export results to CSV? (Y/N)"
        if ($ExportChoice -match "^[Yy]") {
            $ExportPath = "C:\Temp\MailboxForwards_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            $Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            Write-Host "✓ Results exported to: $ExportPath" -ForegroundColor Green
        }
    }
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    exit 1
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host ""
    Write-Host "Disconnected from Exchange Online." -ForegroundColor DarkGray
}

<# Usage:
Simply run the script and follow the interactive prompts:
.\RetrieveMailboxForwards.ps1

Or specify your admin account:
.\RetrieveMailboxForwards.ps1 -AdminUPN "admin@company.com"
#>