<#
.SYNOPSIS
    Retrieves all user permissions for a SharePoint site
.DESCRIPTION
    Audits SharePoint site permissions including users, groups, and permission levels
    Outputs results to console and CSV file
.NOTES
    Requires PnP.PowerShell module
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "C:\Temp\SharePointPermissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " SHAREPOINT SITE PERMISSIONS AUDIT" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Install and import PnP PowerShell module
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Yellow
    Install-Module PnP.PowerShell -Force -Scope CurrentUser -AllowClobber
}
Import-Module PnP.PowerShell -Force

# Get site URL if not provided
if (-not $SiteUrl) {
    Write-Host "Enter SharePoint site URL" -ForegroundColor Cyan
    Write-Host "Examples:" -ForegroundColor Gray
    Write-Host "  https://contoso.sharepoint.com/sites/TeamSite" -ForegroundColor Gray
    Write-Host "  https://contoso.sharepoint.com" -ForegroundColor Gray
    Write-Host ""
    $SiteUrl = Read-Host "Site URL"
}

# Validate URL
if ($SiteUrl -notmatch '^https://.*\.sharepoint\.com') {
    Write-Host "ERROR: Invalid SharePoint URL format" -ForegroundColor Red
    Write-Host "URL must start with https:// and contain .sharepoint.com" -ForegroundColor Yellow
    exit 1
}

# Connect to SharePoint site
Write-Host ""
Write-Host "Connecting to SharePoint site..." -ForegroundColor Cyan
Write-Host "URL: $SiteUrl" -ForegroundColor White

try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId "a5be1f75-9cc4-46b8-9787-2866e7c3c59c" -ErrorAction Stop
    Write-Host "Connected successfully." -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to connect to SharePoint site" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Retrieving site information..." -ForegroundColor Cyan

try {
    $Site = Get-PnPSite -Includes Owner
    $Web = Get-PnPWeb
    
    Write-Host "Site Title: $($Web.Title)" -ForegroundColor White
    Write-Host "Site URL: $($Web.Url)" -ForegroundColor White
    Write-Host ""
}
catch {
    Write-Host "ERROR: Could not retrieve site information" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Disconnect-PnPOnline
    exit 1
}

# Collection to store all permissions
$AllPermissions = @()

Write-Host "Gathering permissions data..." -ForegroundColor Cyan
Write-Host ""

# ============================================
# SITE COLLECTION ADMINISTRATORS
# ============================================
Write-Host "[1/4] Site Collection Administrators..." -ForegroundColor Yellow

try {
    $SiteAdmins = Get-PnPSiteCollectionAdmin
    
    foreach ($Admin in $SiteAdmins) {
        $AllPermissions += [PSCustomObject]@{
            Type = "Site Collection Admin"
            Name = $Admin.Title
            LoginName = $Admin.LoginName
            Email = $Admin.Email
            PermissionLevel = "Site Collection Administrator"
            GrantedThrough = "Direct"
            Location = "Site Collection"
        }
        Write-Host "  + $($Admin.Title)" -ForegroundColor Gray
    }
    
    Write-Host "  Found: $($SiteAdmins.Count) administrators" -ForegroundColor Green
}
catch {
    Write-Host "  Warning: Could not retrieve site collection admins: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host ""

# ============================================
# SITE GROUPS AND MEMBERS
# ============================================
Write-Host "[2/4] SharePoint Groups..." -ForegroundColor Yellow

try {
    $Groups = Get-PnPGroup
    
    foreach ($Group in $Groups) {
        Write-Host "  Processing group: $($Group.Title)" -ForegroundColor Gray
        
        try {
            # Get group members
            $Members = Get-PnPGroupMember -Identity $Group.Title
            
            # Get group permissions
            $GroupPermissions = Get-PnPGroupPermissions -Identity $Group.Title -ErrorAction SilentlyContinue
            $PermissionLevels = if ($GroupPermissions) { 
                ($GroupPermissions.Name -join ", ") 
            } else { 
                "Unknown" 
            }
            
            foreach ($Member in $Members) {
                $AllPermissions += [PSCustomObject]@{
                    Type = "User"
                    Name = $Member.Title
                    LoginName = $Member.LoginName
                    Email = $Member.Email
                    PermissionLevel = $PermissionLevels
                    GrantedThrough = "Group: $($Group.Title)"
                    Location = "Site"
                }
            }
            
            Write-Host "    Members: $($Members.Count)" -ForegroundColor DarkGray
        }
        catch {
            Write-Host "    Warning: Could not retrieve members for $($Group.Title)" -ForegroundColor Yellow
        }
    }
    
    Write-Host "  Found: $($Groups.Count) groups" -ForegroundColor Green
}
catch {
    Write-Host "  Error retrieving groups: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""

# ============================================
# DIRECT USER PERMISSIONS (not in groups)
# ============================================
Write-Host "[3/4] Direct User Permissions..." -ForegroundColor Yellow

try {
    $RoleAssignments = Get-PnPProperty -ClientObject $Web -Property RoleAssignments
    
    $DirectUserCount = 0
    foreach ($RoleAssignment in $RoleAssignments) {
        Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
        
        # Check if it's a user (not a group)
        if ($RoleAssignment.Member.PrincipalType -eq "User") {
            $User = $RoleAssignment.Member
            $Permissions = ($RoleAssignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name) -join ", "
            
            $AllPermissions += [PSCustomObject]@{
                Type = "User"
                Name = $User.Title
                LoginName = $User.LoginName
                Email = $User.Email
                PermissionLevel = $Permissions
                GrantedThrough = "Direct"
                Location = "Site"
            }
            
            $DirectUserCount++
            Write-Host "  + $($User.Title)" -ForegroundColor Gray
        }
    }
    
    Write-Host "  Found: $DirectUserCount direct user permissions" -ForegroundColor Green
}
catch {
    Write-Host "  Warning: Could not retrieve direct permissions: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host ""

# ============================================
# SUMMARY AND EXPORT
# ============================================
Write-Host "[4/4] Generating Report..." -ForegroundColor Yellow

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " PERMISSIONS SUMMARY" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

$TotalUsers = ($AllPermissions | Where-Object { $_.Type -eq "User" }).Count
$TotalAdmins = ($AllPermissions | Where-Object { $_.Type -eq "Site Collection Admin" }).Count
$UniqueUsers = ($AllPermissions | Select-Object -Property LoginName -Unique).Count

Write-Host "Site: $($Web.Title)" -ForegroundColor White
Write-Host "Total permission entries: $($AllPermissions.Count)" -ForegroundColor White
Write-Host "Unique users: $UniqueUsers" -ForegroundColor White
Write-Host "Site Collection Admins: $TotalAdmins" -ForegroundColor White
Write-Host ""

# Display sample (first 15 entries)
Write-Host "Sample Permissions (first 15):" -ForegroundColor Cyan
$AllPermissions | Select-Object -First 15 | Format-Table Name, PermissionLevel, GrantedThrough -AutoSize

# Export to CSV
Write-Host ""
Write-Host "Exporting to CSV..." -ForegroundColor Cyan

if (-not (Test-Path "C:\Temp")) {
    New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null
}

$AllPermissions | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Export complete: $OutputPath" -ForegroundColor Green
Write-Host ""

# Permission level breakdown
Write-Host "Permission Level Breakdown:" -ForegroundColor Cyan
$AllPermissions | Group-Object PermissionLevel | Sort-Object Count -Descending | ForEach-Object {
    Write-Host "  $($_.Name): $($_.Count)" -ForegroundColor White
}

Write-Host ""

# Disconnect
Disconnect-PnPOnline
Write-Host "Disconnected from SharePoint." -ForegroundColor DarkGray
Write-Host "Operation complete." -ForegroundColor Green

<#
.USAGE EXAMPLES

Interactive mode (prompts for site URL):
.\SharepointPermissions.ps1

With site URL provided:
.\SharepointPermissions.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/TeamSite"

With custom output path:
.\SharepointPermissions.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/TeamSite" -OutputPath "C:\Reports\SitePermissions.csv"
#>