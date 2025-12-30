<#
.SYNOPSIS
Extracts all groups a user belongs to - uses browser-based global admin authentication.

.DESCRIPTION
Opens a browser window for you to sign in with global admin credentials.
Then queries Microsoft Graph to retrieve all groups the user belongs to.

.PARAMETER UserPrincipalName
The email/UPN of the user (e.g., michael.weir@domain.com).

.PARAMETER TenantId
Your Azure tenant ID or domain name (e.g., contoso.onmicrosoft.com).

.PARAMETER OutputPath
Path where the CSV report will be saved.

.EXAMPLE
.\Get-UserSharePointGroups.ps1 -UserPrincipalName 'michael.weir@domain.com' -TenantId 'cpcengineering.onmicrosoft.com'
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, HelpMessage="User email/UPN")]
    [string]$UserPrincipalName,

    [Parameter(Mandatory=$true, HelpMessage="Tenant ID or domain")]
    [string]$TenantId,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\UserGroups_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Azure AD app registration (standard public client for PowerShell)
$clientId = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"  # Azure CLI app (works for device+browser flow)
$redirectUri = "http://localhost"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SharePoint Groups Extraction Tool" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Opening browser for authentication..." -ForegroundColor Yellow
Write-Host "Please sign in with your GLOBAL ADMIN credentials." -ForegroundColor Yellow
Write-Host ""

# Step 1: Authorization endpoint (browser-based)
$authUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/authorize"
$authParams = @{
    client_id     = $clientId
    response_type = "code"
    scope         = "https://graph.microsoft.com/.default"
    redirect_uri  = $redirectUri
    prompt        = "login"  # Force login with admin creds
}

$authUri = $authUrl + "?" + (($authParams.GetEnumerator() | ForEach-Object { "$($_.Key)=$([uri]::EscapeDataString($_.Value))" }) -join "&")

# Open browser
Write-Host "Opening: $authUrl" -ForegroundColor Cyan
Start-Process $authUri

Write-Host ""
Write-Host "Waiting for browser authentication..." -ForegroundColor Yellow
Write-Host "If browser doesn't open, visit:" -ForegroundColor Gray
Write-Host $authUri -ForegroundColor Gray
Write-Host ""

# Step 2: Get auth code from user
Write-Host "After authentication, you'll be redirected to http://localhost with a 'code' parameter." -ForegroundColor Yellow
$authCode = Read-Host "Paste the authorization code from the URL"

if (-not $authCode) {
    Write-Error "No authorization code provided."
    exit 1
}

# Step 3: Exchange code for token
Write-Host "Exchanging code for access token..." -ForegroundColor Cyan

$tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$tokenBody = @{
    grant_type    = "authorization_code"
    client_id     = $clientId
    code          = $authCode
    redirect_uri  = $redirectUri
    scope         = "https://graph.microsoft.com/.default"
}

try {
    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $tokenBody -ErrorAction Stop
    $accessToken = $tokenResponse.access_token
    Write-Host "Authentication successful!" -ForegroundColor Green
}
catch {
    Write-Error "Failed to get access token: $_"
    exit 1
}

# Step 4: Query Graph API
$headers = @{
    Authorization = "Bearer $accessToken"
    "Content-Type" = "application/json"
}

Write-Host ""
Write-Host "Searching for user: $UserPrincipalName" -ForegroundColor Cyan

# Get user
$userUri = "https://graph.microsoft.com/v1.0/users?`$filter=userPrincipalName eq '$UserPrincipalName'"
try {
    $userResponse = Invoke-RestMethod -Method Get -Uri $userUri -Headers $headers -ErrorAction Stop
    
    if ($userResponse.value.Count -eq 0) {
        Write-Error "User not found: $UserPrincipalName"
        exit 1
    }
    
    $user = $userResponse.value[0]
    Write-Host "Found: $($user.displayName) ($($user.userPrincipalName))" -ForegroundColor Green
}
catch {
    Write-Error "Error searching for user: $_"
    exit 1
}

# Get groups
Write-Host "Retrieving group memberships..." -ForegroundColor Cyan

$results = @()

try {
    $groupsUri = "https://graph.microsoft.com/v1.0/users/$($user.id)/memberOf?`$top=999"
    $groupsResponse = Invoke-RestMethod -Method Get -Uri $groupsUri -Headers $headers -ErrorAction Stop
    
    $groups = $groupsResponse.value
    
    # Handle pagination
    while ($groupsResponse.'@odata.nextLink') {
        $groupsResponse = Invoke-RestMethod -Method Get -Uri $groupsResponse.'@odata.nextLink' -Headers $headers -ErrorAction Stop
        $groups += $groupsResponse.value
    }
    
    Write-Host "Found $($groups.Count) group(s)." -ForegroundColor Green
    Write-Host ""
    
    foreach ($group in $groups) {
        $groupType = if ($group.'@odata.type' -like '*group*') { "Group" } else { $group.'@odata.type' -replace ".*\." }
        $results += [PSCustomObject]@{
            GroupName        = $group.displayName
            GroupId          = $group.id
            GroupType        = $groupType
            Mail             = $group.mail
            Description      = $group.description
            CreatedDate      = $group.createdDateTime
            ExtractionDate   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        
        Write-Host "  ✓ $($group.displayName) ($groupType)" -ForegroundColor Green
    }
}
catch {
    Write-Error "Error retrieving groups: $_"
    exit 1
}

# Export
Write-Host ""
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Force
    Write-Host "✓ Exported to: $OutputPath" -ForegroundColor Green
    Write-Host "✓ Total groups: $($results.Count)" -ForegroundColor Green
}
else {
    Write-Host "No groups found." -ForegroundColor Yellow
}

Write-Host "Done." -ForegroundColor Cyan
