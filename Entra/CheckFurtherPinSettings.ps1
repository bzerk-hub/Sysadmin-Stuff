param(
    [Parameter(Mandatory = $false)]
    [string]$AdminUPN
)

$ErrorActionPreference = 'Stop'

try {
    # Connect to Microsoft Graph
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host " ADDITIONAL PIN POLICY CHECKS" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    
    $Scopes = @(
        "DeviceManagementApps.Read.All",
        "DeviceManagementConfiguration.Read.All",
        "Policy.Read.All",
        "Directory.Read.All"
    )
    
    if ($AdminUPN) {
        Connect-MgGraph -Scopes $Scopes -AccountId $AdminUPN -NoWelcome
    } else {
        Connect-MgGraph -Scopes $Scopes -NoWelcome
    }
    
    Write-Host "Connected." -ForegroundColor Green
    Write-Host ""
    
    $AllResults = @()
    
    # ============================================
    # 1. CHECK ASSIGNED POLICIES FOR SPECIFIC USER
    # ============================================
    Write-Host "[1/8] Enter a test user's email to check their assigned policies:" -ForegroundColor Cyan
    $TestUser = Read-Host "User email (or press Enter to skip)"
    
    if ($TestUser) {
        Write-Host "Checking policies assigned to $TestUser..." -ForegroundColor Yellow
        
        try {
            # Get user's group memberships
            $User = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$TestUser" -Method GET
            $Groups = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$($User.id)/memberOf" -Method GET
            
            Write-Host "  User is member of $($Groups.value.Count) groups" -ForegroundColor DarkGray
            
            # Check iOS policies assigned to user/groups
            $iOSPolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/iosManagedAppProtections" -Method GET
            foreach ($Policy in $iOSPolicies.value) {
                $Assignments = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/iosManagedAppProtections/$($Policy.id)/assignments" -Method GET -ErrorAction SilentlyContinue
                
                $IsAssigned = $false
                foreach ($Assignment in $Assignments.value) {
                    if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
                        $IsAssigned = $true
                        break
                    }
                    if ($Assignment.target.groupId -in $Groups.value.id) {
                        $IsAssigned = $true
                        break
                    }
                }
                
                if ($IsAssigned) {
                    $AllResults += [PSCustomObject]@{
                        Source              = "User-Assigned iOS Policy"
                        PolicyName          = $Policy.displayName
                        Platform            = "iOS"
                        MinimumPINLength    = $Policy.minimumPinLength
                        PINCharacterSet     = $Policy.pinCharacterSet
                        AssignedTo          = $TestUser
                        PolicyId            = $Policy.id
                    }
                }
            }
            
            # Check Android policies
            $AndroidPolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections" -Method GET
            foreach ($Policy in $AndroidPolicies.value) {
                $Assignments = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections/$($Policy.id)/assignments" -Method GET -ErrorAction SilentlyContinue
                
                $IsAssigned = $false
                foreach ($Assignment in $Assignments.value) {
                    if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
                        $IsAssigned = $true
                        break
                    }
                    if ($Assignment.target.groupId -in $Groups.value.id) {
                        $IsAssigned = $true
                        break
                    }
                }
                
                if ($IsAssigned) {
                    $AllResults += [PSCustomObject]@{
                        Source              = "User-Assigned Android Policy"
                        PolicyName          = $Policy.displayName
                        Platform            = "Android"
                        MinimumPINLength    = $Policy.minimumPinLength
                        PINCharacterSet     = $Policy.pinCharacterSet
                        AssignedTo          = $TestUser
                        PolicyId            = $Policy.id
                    }
                }
            }
        }
        catch {
            Write-Host "  ✗ Could not check user assignments: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    # ============================================
    # 2. CHECK TENANT DEFAULT APP PROTECTION SETTINGS
    # ============================================
    Write-Host "[2/8] Checking Tenant Default MAM Settings..." -ForegroundColor Cyan
    
    try {
        $MAMSettings = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/managedAppPolicies?`$filter=isof('microsoft.graph.targetedManagedAppConfiguration')" -Method GET
        foreach ($Setting in $MAMSettings.value) {
            if ($Setting.minimumPinLength) {
                $AllResults += [PSCustomObject]@{
                    Source              = "Default MAM Setting"
                    PolicyName          = $Setting.displayName
                    Platform            = "Cross-Platform"
                    MinimumPINLength    = $Setting.minimumPinLength
                    PINCharacterSet     = $Setting.pinCharacterSet
                    AssignedTo          = "Tenant Default"
                    PolicyId            = $Setting.id
                }
            }
        }
        Write-Host "  ✓ Checked default MAM settings" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # ============================================
    # 3. CHECK WINDOWS INFORMATION PROTECTION (WIP)
    # ============================================
    Write-Host "[3/8] Checking Windows Information Protection Policies..." -ForegroundColor Cyan
    
    try {
        $WIPPolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/windowsInformationProtectionPolicies" -Method GET
        foreach ($Policy in $WIPPolicies.value) {
            if ($Policy.pinMinimumLength) {
                $AllResults += [PSCustomObject]@{
                    Source              = "Windows Information Protection"
                    PolicyName          = $Policy.displayName
                    Platform            = "Windows"
                    MinimumPINLength    = $Policy.pinMinimumLength
                    PINCharacterSet     = $Policy.pinCharacterSet
                    AssignedTo          = "See Assignments"
                    PolicyId            = $Policy.id
                }
            }
        }
        Write-Host "  ✓ Found $($WIPPolicies.value.Count) WIP policies" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # ============================================
    # 4. CHECK MDM AUTHORITY SETTINGS
    # ============================================
    Write-Host "[4/8] Checking MDM Authority Settings..." -ForegroundColor Cyan
    
    try {
        $MDMAuthority = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement" -Method GET
        Write-Host "  MDM Authority: $($MDMAuthority.intuneAccountId)" -ForegroundColor DarkGray
        Write-Host "  Intune Brand: $($MDMAuthority.intuneBrand.displayName)" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "  ✗ Could not retrieve MDM settings" -ForegroundColor Red
    }
    
    # ============================================
    # 5. CHECK APP-SPECIFIC CONFIGURATIONS
    # ============================================
    Write-Host "[5/8] Checking Managed App Configurations (per-app settings)..." -ForegroundColor Cyan
    
    try {
        $ManagedAppConfigs = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/targetedManagedAppConfigurations" -Method GET
        foreach ($Config in $ManagedAppConfigs.value) {
            $Details = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/targetedManagedAppConfigurations/$($Config.id)" -Method GET
            
            if ($Details.customSettings) {
                foreach ($Setting in $Details.customSettings) {
                    if ($Setting.name -match "pin|password" -and $Setting.value -match "8") {
                        $AllResults += [PSCustomObject]@{
                            Source              = "Managed App Config"
                            PolicyName          = $Config.displayName
                            Platform            = "App-Specific"
                            MinimumPINLength    = "Custom: $($Setting.value)"
                            PINCharacterSet     = "See Config"
                            AssignedTo          = "See Assignments"
                            PolicyId            = $Config.id
                        }
                    }
                }
            }
        }
        Write-Host "  ✓ Found $($ManagedAppConfigs.value.Count) managed app configs" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # ============================================
    # 6. CHECK LEGACY INTUNE POLICIES
    # ============================================
    Write-Host "[6/8] Checking Legacy Intune Policies..." -ForegroundColor Cyan
    
    try {
        $LegacyPolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts" -Method GET
        Write-Host "  ✓ Found $($LegacyPolicies.value.Count) PowerShell scripts" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # ============================================
    # 7. CHECK AUTHENTICATION METHODS POLICY
    # ============================================
    Write-Host "[7/8] Checking Authentication Methods Policy..." -ForegroundColor Cyan
    
    try {
        $AuthMethods = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/policies/authenticationMethodsPolicy" -Method GET
        if ($AuthMethods.authenticationMethodConfigurations) {
            Write-Host "  Found authentication method configurations" -ForegroundColor DarkGray
            
            foreach ($Method in $AuthMethods.authenticationMethodConfigurations) {
                if ($Method.'@odata.type' -match "microsoft" -and $Method.state -eq "enabled") {
                    Write-Host "    - $($Method.'@odata.type') is enabled" -ForegroundColor DarkGray
                }
            }
        }
    }
    catch {
        Write-Host "  ✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # ============================================
    # 8. CHECK MICROSOFT MANAGED APP (INTUNE APP SDK)
    # ============================================
    Write-Host "[8/8] Checking for Intune App SDK Defaults..." -ForegroundColor Cyan
    
    Write-Host "  Note: Some apps (Outlook, Teams, OneDrive) may have SDK defaults" -ForegroundColor Yellow
    Write-Host "  that override policy settings. Check in-app settings." -ForegroundColor Yellow
    
    # ============================================
    # DISPLAY RESULTS
    # ============================================
    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host " ADDITIONAL FINDINGS" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    
    if ($AllResults.Count -eq 0) {
        Write-Host "⚠ No additional policies found enforcing 8-digit PIN" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "POSSIBLE CAUSES:" -ForegroundColor Red
        Write-Host "  1. App SDK Default - Some Microsoft apps default to 8 digits" -ForegroundColor White
        Write-Host "  2. Device OS Requirement - iOS/Android may enforce higher minimums" -ForegroundColor White
        Write-Host "  3. Multiple Policies Conflict - Highest value wins" -ForegroundColor White
        Write-Host "  4. Cached Policy - User needs to unenroll/re-enroll device" -ForegroundColor White
        Write-Host "  5. Third-party MDM/MAM - Another management tool is active" -ForegroundColor White
        Write-Host ""
        Write-Host "TROUBLESHOOTING STEPS:" -ForegroundColor Cyan
        Write-Host "  1. In Intune Portal > Apps > App protection policies > [Policy] > Properties" -ForegroundColor White
        Write-Host "     Check 'Data protection' section for PIN requirements" -ForegroundColor White
        Write-Host "  2. Have user go to Settings > General > Profiles & Device Management" -ForegroundColor White
        Write-Host "     Check what profiles are installed" -ForegroundColor White
        Write-Host "  3. Check the specific app's settings (e.g., Outlook > Settings > PIN)" -ForegroundColor White
        Write-Host "  4. Review Intune logs: Devices > [User Device] > Device configuration" -ForegroundColor White
    } else {
        Write-Host "Found $($AllResults.Count) additional policy/policies:" -ForegroundColor Yellow
        Write-Host ""
        $AllResults | Format-Table -AutoSize
        
        # Export
        Write-Host ""
        $ExportChoice = Read-Host "Export results to CSV? (Y/N)"
        if ($ExportChoice -match "^[Yy]") {
            if (-not (Test-Path "C:\Temp")) {
                New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null
            }
            $ExportPath = "C:\Temp\AdditionalPINCheck_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            $AllResults | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            Write-Host "✓ Results exported to: $ExportPath" -ForegroundColor Green
        }
    }
    
    Write-Host ""
    Write-Host "Done." -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}

<# Usage:
.\DeepPINCheck.ps1
.\DeepPINCheck.ps1 -AdminUPN "admin@company.com"
#>