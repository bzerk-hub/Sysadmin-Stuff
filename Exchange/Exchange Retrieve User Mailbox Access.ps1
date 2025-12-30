param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter the email or username of the user whose mailbox access you want to check")]
    [Alias("User", "Email", "TargetUser")]
    [string]$UserIdentity,
    
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
    Write-Host " Exchange Mailbox Access Report" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    if ($AdminUPN) {
        Write-Host "Using admin account: $AdminUPN" -ForegroundColor DarkGray
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false
    } else {
        Write-Host "(You will be prompted to sign in with your admin account)" -ForegroundColor DarkGray
        Connect-ExchangeOnline -ShowBanner:$false
    }
    Write-Host "Connected." -ForegroundColor Green
    Write-Host ""
    
    # Resolve the user being checked (with ambiguity handling)
    Write-Host "Looking up TARGET USER: $UserIdentity..." -ForegroundColor Cyan
    
    # Try exact match by primary SMTP first
    $User = Get-EXOMailbox -Identity $UserIdentity -ErrorAction SilentlyContinue
    
    # If not found or ambiguous, search by filter
    if (-not $User) {
        Write-Host "Exact match not found. Searching..." -ForegroundColor Yellow
        $Candidates = @()
        
        # Search mailboxes
        $Candidates += Get-EXOMailbox -Filter "PrimarySmtpAddress -eq '$UserIdentity' -or UserPrincipalName -eq '$UserIdentity' -or Alias -eq '$UserIdentity'" -ErrorAction SilentlyContinue
        
        if ($Candidates.Count -eq 0) {
            # Try broader search
            $Candidates = Get-EXOMailbox -Filter "DisplayName -like '*$UserIdentity*' -or Alias -like '*$UserIdentity*'" -ErrorAction SilentlyContinue
        }
        
        if ($Candidates.Count -eq 0) {
            throw "No mailbox found matching '$UserIdentity'. Please use the full email address (e.g., user@domain.com)."
        }
        
        if ($Candidates.Count -gt 1) {
            Write-Host ""
            Write-Host "Multiple mailboxes found matching '$UserIdentity':" -ForegroundColor Yellow
            Write-Host ""
            $Candidates | Format-Table -Property DisplayName, PrimarySmtpAddress, RecipientTypeDetails -AutoSize
            Write-Host ""
            throw "Ambiguous identity. Please specify the exact email address from the list above."
        }
        
        $User = $Candidates[0]
    }
    
    Write-Host "✓ Target user found: $($User.DisplayName) ($($User.PrimarySmtpAddress))" -ForegroundColor Green
    Write-Host ""
    Write-Host "Checking which mailboxes this user has access to..." -ForegroundColor Yellow
    Write-Host ""
    
    # Get all mailboxes in the organization
    Write-Host "Scanning all mailboxes for permissions (this may take 5-10 minutes)..." -ForegroundColor Yellow
    $AllMailboxes = Get-EXOMailbox -ResultSize Unlimited
    
    $Results = @()
    $Counter = 0
    $Total = $AllMailboxes.Count
    
    foreach ($Mailbox in $AllMailboxes) {
        $Counter++
        Write-Progress -Activity "Checking mailbox permissions for $($User.DisplayName)" -Status "Processing mailbox $Counter of $Total" -PercentComplete (($Counter / $Total) * 100)
        
        # Check Full Access permissions
        $FullAccess = Get-EXOMailboxPermission -Identity $Mailbox.PrimarySmtpAddress | 
            Where-Object { $_.User -eq $User.PrimarySmtpAddress -and $_.AccessRights -contains "FullAccess" -and $_.IsInherited -eq $false }
        
        # Check Send As permissions
        $SendAs = Get-EXORecipientPermission -Identity $Mailbox.PrimarySmtpAddress | 
            Where-Object { $_.Trustee -eq $User.PrimarySmtpAddress -and $_.AccessRights -contains "SendAs" }
        
        # Check Send on Behalf permissions
        $SendOnBehalf = if ($Mailbox.GrantSendOnBehalfTo -contains $User.PrimarySmtpAddress) { $true } else { $false }
        
        # Build result object
        if ($FullAccess -or $SendAs -or $SendOnBehalf) {
            $Permissions = @()
            if ($FullAccess) { $Permissions += "Full Access" }
            if ($SendAs) { $Permissions += "Send As" }
            if ($SendOnBehalf) { $Permissions += "Send on Behalf" }
            
            $Results += [PSCustomObject]@{
                Mailbox           = $Mailbox.DisplayName
                PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
                MailboxType       = $Mailbox.RecipientTypeDetails
                Permissions       = $Permissions -join ", "
            }
        }
    }
    
    Write-Progress -Activity "Checking mailbox permissions" -Completed
    
    # Display results
    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host " RESULTS" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    
    if ($Results.Count -eq 0) {
        Write-Host "✗ No mailbox permissions found for $($User.DisplayName)." -ForegroundColor Yellow
        Write-Host "  This user does not have Full Access, Send As, or Send on Behalf permissions to any mailboxes." -ForegroundColor DarkGray
    } else {
        Write-Host "✓ Mailbox Access Report for: $($User.DisplayName) ($($User.PrimarySmtpAddress))" -ForegroundColor Green
        Write-Host "  Total mailboxes with access: $($Results.Count)" -ForegroundColor Cyan
        Write-Host ""
        $Results | Format-Table -AutoSize
        
        # Optional: Export to CSV
        $ExportPath = "C:\Temp\MailboxAccess_$($User.DisplayName -replace '[^a-zA-Z0-9]','_')_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host "✓ Results exported to: $ExportPath" -ForegroundColor Green
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

<# Usage Examples:

Best practice - use full email address:
.\Exchange Retrieve User Mailbox Access.ps1 -UserIdentity "john.doe@company.com"

If ambiguous, script will list matches and ask for exact email:
.\Exchange Retrieve User Mailbox Access.ps1 -UserIdentity "Marketing"

Specify admin account upfront:
.\Exchange Retrieve User Mailbox Access.ps1 -UserIdentity "john.doe@company.com" -AdminUPN "admin@company.com"

#>