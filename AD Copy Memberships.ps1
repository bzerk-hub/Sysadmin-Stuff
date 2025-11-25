param(
    [Parameter(Mandatory = $true)]
    [string]$SourceUser,
    
    [Parameter(Mandatory = $true)]
    [string]$TargetUser
)

$ErrorActionPreference = 'Stop'

try {
    # Import Active Directory module
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Validate source user exists
    Write-Host "Validating source user: $SourceUser..." -ForegroundColor Cyan
    $Source = Get-ADUser -Identity $SourceUser -Properties MemberOf -ErrorAction Stop
    Write-Host "Source user found: $($Source.Name)" -ForegroundColor Green
    
    # Validate target user exists
    Write-Host "Validating target user: $TargetUser..." -ForegroundColor Cyan
    $Target = Get-ADUser -Identity $TargetUser -ErrorAction Stop
    Write-Host "Target user found: $($Target.Name)" -ForegroundColor Green
    
    # Get groups of source user
    $Groups = $Source.MemberOf
    
    if (-not $Groups -or $Groups.Count -eq 0) {
        Write-Host "Source user is not a member of any groups." -ForegroundColor Yellow
        exit 0
    }
    
    Write-Host "`nCopying $($Groups.Count) group memberships..." -ForegroundColor Cyan
    
    $SuccessCount = 0
    $SkippedCount = 0
    $FailedCount = 0
    
    foreach ($GroupDN in $Groups) {
        $Group = Get-ADGroup -Identity $GroupDN
        
        # Check if target is already a member
        $IsMember = Get-ADGroupMember -Identity $Group | Where-Object { $_.SamAccountName -eq $Target.SamAccountName }
        
        if ($IsMember) {
            Write-Host "  [SKIP] $($Group.Name) - Already a member" -ForegroundColor DarkGray
            $SkippedCount++
            continue
        }
        
        try {
            Add-ADGroupMember -Identity $Group -Members $Target -ErrorAction Stop
            Write-Host "  [OK] $($Group.Name)" -ForegroundColor Green
            $SuccessCount++
        } catch {
            Write-Warning "  [FAIL] $($Group.Name): $($_.Exception.Message)"
            $FailedCount++
        }
    }
    
    Write-Host "`nSummary:" -ForegroundColor Cyan
    Write-Host "  Added: $SuccessCount" -ForegroundColor Green
    Write-Host "  Skipped (already member): $SkippedCount" -ForegroundColor DarkGray
    Write-Host "  Failed: $FailedCount" -ForegroundColor $(if ($FailedCount -gt 0) { 'Red' } else { 'Green' })
    Write-Host "Done." -ForegroundColor Green
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

<# Usage:
.\AD Copy Memberships.ps1 -SourceUser "john.doe" -TargetUser "jane.smith"
#>