<#
.SYNOPSIS
    Retrieves detailed permissions for all folders and files in a SharePoint site
.DESCRIPTION
    Audits SharePoint permissions at site, library, folder, and file level
    Shows who has access to specific folders and what permissions they have
.NOTES
    Requires PnP.PowerShell module
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "C:\Temp\SharePointFolderPermissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " SHAREPOINT FOLDER PERMISSIONS AUDIT" -ForegroundColor Cyan
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
    Write-Host ""
    $SiteUrl = Read-Host "Site URL"
}

# Validate URL
if ($SiteUrl -notmatch '^https://.*\.sharepoint\.com') {
    Write-Host "ERROR: Invalid SharePoint URL format" -ForegroundColor Red
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
    Write-Host "ERROR: Failed to connect: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
$Web = Get-PnPWeb
Write-Host "Site: $($Web.Title)" -ForegroundColor White
Write-Host ""

# Collection to store all permissions
$AllPermissions = @()

# Function to get permissions for an item
function Get-ItemPermissions {
    param(
        [Parameter(Mandatory=$true)]
        $Item,
        
        [Parameter(Mandatory=$true)]
        [string]$ItemType,
        
        [Parameter(Mandatory=$true)]
        [string]$LibraryName,
        
        [Parameter(Mandatory=$true)]
        [string]$ItemPath
    )
    
    try {
        # Check if item has unique permissions
        if ($Item.HasUniqueRoleAssignments) {
            $RoleAssignments = Get-PnPProperty -ClientObject $Item -Property RoleAssignments
            
            foreach ($RoleAssignment in $RoleAssignments) {
                Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
                
                $Member = $RoleAssignment.Member
                $Permissions = ($RoleAssignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name) -join ", "
                
                # Check if it's a group or user
                if ($Member.PrincipalType -eq "SharePointGroup") {
                    # Get group members
                    try {
                        $GroupMembers = Get-PnPGroupMember -Identity $Member.LoginName -ErrorAction SilentlyContinue
                        
                        foreach ($GroupMember in $GroupMembers) {
                            $script:AllPermissions += [PSCustomObject]@{
                                Library = $LibraryName
                                ItemType = $ItemType
                                Path = $ItemPath
                                UserName = $GroupMember.Title
                                UserEmail = $GroupMember.Email
                                PermissionLevel = $Permissions
                                GrantedThrough = "Group: $($Member.Title)"
                                HasUniquePermissions = "Yes"
                            }
                        }
                    }
                    catch {
                        $script:AllPermissions += [PSCustomObject]@{
                            Library = $LibraryName
                            ItemType = $ItemType
                            Path = $ItemPath
                            UserName = $Member.Title
                            UserEmail = ""
                            PermissionLevel = $Permissions
                            GrantedThrough = "Group (members not retrieved)"
                            HasUniquePermissions = "Yes"
                        }
                    }
                }
                else {
                    $script:AllPermissions += [PSCustomObject]@{
                        Library = $LibraryName
                        ItemType = $ItemType
                        Path = $ItemPath
                        UserName = $Member.Title
                        UserEmail = $Member.Email
                        PermissionLevel = $Permissions
                        GrantedThrough = "Direct"
                        HasUniquePermissions = "Yes"
                    }
                }
            }
        }
        else {
            # Inherits permissions - note this but don't enumerate all inherited permissions
            $script:AllPermissions += [PSCustomObject]@{
                Library = $LibraryName
                ItemType = $ItemType
                Path = $ItemPath
                UserName = "(Inherited)"
                UserEmail = ""
                PermissionLevel = "Inherited from parent"
                GrantedThrough = "Inheritance"
                HasUniquePermissions = "No"
            }
        }
    }
    catch {
        Write-Host "    Warning: Could not retrieve permissions for $ItemPath" -ForegroundColor Yellow
    }
}

# Get all document libraries
Write-Host "Retrieving document libraries..." -ForegroundColor Cyan
$Libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }

Write-Host "Found $($Libraries.Count) document libraries" -ForegroundColor Green
Write-Host ""

$LibraryCounter = 0
foreach ($Library in $Libraries) {
    $LibraryCounter++
    Write-Host "[$LibraryCounter/$($Libraries.Count)] Processing: $($Library.Title)" -ForegroundColor Yellow
    
    try {
        # Get all folders and files in the library
        $Items = Get-PnPListItem -List $Library.Title -PageSize 500
        
        $ItemCounter = 0
        foreach ($Item in $Items) {
            $ItemCounter++
            
            if ($ItemCounter % 50 -eq 0) {
                Write-Host "  Processing item $ItemCounter of $($Items.Count)..." -ForegroundColor DarkGray
            }
            
            $ItemType = if ($Item.FileSystemObjectType -eq "Folder") { "Folder" } else { "File" }
            $ItemPath = $Item.FieldValues.FileRef
            
            Get-ItemPermissions -Item $Item -ItemType $ItemType -LibraryName $Library.Title -ItemPath $ItemPath
        }
        
        Write-Host "  Completed: $($Items.Count) items processed" -ForegroundColor Green
    }
    catch {
        Write-Host "  Error processing library: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host ""
}

# Summary and Export
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " GENERATING REPORT" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Total permission entries: $($AllPermissions.Count)" -ForegroundColor White
Write-Host "Items with unique permissions: $(($AllPermissions | Where-Object { $_.HasUniquePermissions -eq 'Yes' }).Count)" -ForegroundColor White
Write-Host "Items with inherited permissions: $(($AllPermissions | Where-Object { $_.HasUniquePermissions -eq 'No' }).Count)" -ForegroundColor White
Write-Host ""

# Display sample
Write-Host "Sample Results (first 20):" -ForegroundColor Cyan
$AllPermissions | Select-Object -First 20 | Format-Table Library, ItemType, Path, UserName, PermissionLevel -AutoSize -Wrap

# Export to CSV
Write-Host ""
Write-Host "Exporting to CSV..." -ForegroundColor Cyan

if (-not (Test-Path "C:\Temp")) {
    New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null
}

$AllPermissions | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Export complete: $OutputPath" -ForegroundColor Green
Write-Host ""
Write-Host "The CSV file shows:" -ForegroundColor Cyan
Write-Host "  - Library: Which document library the item is in" -ForegroundColor White
Write-Host "  - ItemType: Whether it's a Folder or File" -ForegroundColor White
Write-Host "  - Path: The full path to the item" -ForegroundColor White
Write-Host "  - UserName: Who has access" -ForegroundColor White
Write-Host "  - PermissionLevel: What permissions they have (Read, Edit, Full Control, etc.)" -ForegroundColor White
Write-Host "  - GrantedThrough: How they got access (Direct, Group, or Inheritance)" -ForegroundColor White
Write-Host ""

Disconnect-PnPOnline
Write-Host "Disconnected from SharePoint." -ForegroundColor DarkGray
Write-Host "Operation complete." -ForegroundColor Green