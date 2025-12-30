param(
    [Parameter(Mandatory = $false)]
    [string]$AdminUPN
)

$ErrorActionPreference = 'Stop'

try {
    # Install/Import Microsoft Graph modules
    $RequiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.DeviceManagement',
        'Microsoft.Graph.Identity.SignIns',
        'Microsoft.Graph.Beta.DeviceManagement'
    )
    
    foreach ($Module in $RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $Module)) {
            Write-Host "Installing $Module..." -ForegroundColor Yellow
            Install-Module $Module -Force -Scope CurrentUser -AllowClobber
        }
        Import-Module $Module -Force
    }
    
    # Connect to Microsoft Graph with all required scopes
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host " COMPREHENSIVE PIN POLICY AUDIT" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Write-Host "NOTE: You need Global Admin or Intune Administrator role" -ForegroundColor Yellow
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
    
    # Verify connection and permissions
    $Context = Get-MgContext
    Write-Host "Connected as: $($Context.Account)" -ForegroundColor Green
    Write-Host "Tenant: $($Context.TenantId)" -ForegroundColor DarkGray
    Write-Host ""
    
    $AllResults = @()
    
    # ============================================
    # 1. APP PROTECTION POLICIES
    # ============================================
    Write-Host "[1/7] Checking App Protection Policies..." -ForegroundColor Cyan
    
    try {
        # iOS
        $iOSPolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/iosManagedAppProtections" -Method GET -ErrorAction Stop
        foreach ($Policy in $iOSPolicies.value) {
            $AllResults += [PSCustomObject]@{
                Source              = "App Protection Policy"
                PolicyName          = $Policy.displayName
                Platform            = "iOS"
                PINRequired         = $Policy.pinRequired
                MinimumPINLength    = $Policy.minimumPinLength
                PINCharacterSet     = $Policy.pinCharacterSet
                SimplePINBlocked    = $Policy.simplePinBlocked
                PolicyId            = $Policy.id
                PolicyType          = "iosManagedAppProtection"
            }
        }
        
        # Android
        $AndroidPolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections" -Method GET -ErrorAction Stop
        foreach ($Policy in $AndroidPolicies.value) {
            $AllResults += [PSCustomObject]@{
                Source              = "App Protection Policy"
                PolicyName          = $Policy.displayName
                Platform            = "Android"
                PINRequired         = $Policy.pinRequired
                MinimumPINLength    = $Policy.minimumPinLength
                PINCharacterSet     = $Policy.pinCharacterSet
                SimplePINBlocked    = $Policy.simplePinBlocked
                PolicyId            = $Policy.id
                PolicyType          = "androidManagedAppProtection"
            }
        }
        
        # Windows
        $WindowsPolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/windowsManagedAppProtections" -Method GET -ErrorAction Stop
        foreach ($Policy in $WindowsPolicies.value) {
            $AllResults += [PSCustomObject]@{
                Source              = "App Protection Policy"
                PolicyName          = $Policy.displayName
                Platform            = "Windows"
                PINRequired         = $Policy.pinRequired
                MinimumPINLength    = $Policy.minimumPinLength
                PINCharacterSet     = $Policy.pinCharacterSet
                SimplePINBlocked    = $Policy.simplePinBlocked
                PolicyId            = $Policy.id
                PolicyType          = "windowsManagedAppProtection"
            }
        }
        
        Write-Host "  ✓ Found $($iOSPolicies.value.Count) iOS, $($AndroidPolicies.value.Count) Android, $($WindowsPolicies.value.Count) Windows policies" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Required permission: DeviceManagementApps.Read.All" -ForegroundColor Yellow
    }
    
    # ============================================
    # 2. DEVICE COMPLIANCE POLICIES
    # ============================================
    Write-Host "[2/7] Checking Device Compliance Policies..." -ForegroundColor Cyan
    
    try {
        $CompliancePolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies" -Method GET -ErrorAction Stop
        foreach ($Policy in $CompliancePolicies.value) {
            $Details = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($Policy.id)" -Method GET -ErrorAction Stop
            
            $PINLength = $null
            if ($Details.passwordMinimumLength) { $PINLength = $Details.passwordMinimumLength }
            if ($Details.passcodeMinimumLength) { $PINLength = $Details.passcodeMinimumLength }
            
            if ($PINLength) {
                $AllResults += [PSCustomObject]@{
                    Source              = "Device Compliance Policy"
                    PolicyName          = $Policy.displayName
                    Platform            = $Policy.'@odata.type' -replace '#microsoft.graph.',''
                    PINRequired         = $Details.passwordRequired
                    MinimumPINLength    = $PINLength
                    PINCharacterSet     = $Details.passwordRequiredType
                    SimplePINBlocked    = $Details.simplePasswordBlocked
                    PolicyId            = $Policy.id
                    PolicyType          = "CompliancePolicy"
                }
            }
        }
        Write-Host "  ✓ Found $($CompliancePolicies.value.Count) compliance policies" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Required permission: DeviceManagementConfiguration.Read.All" -ForegroundColor Yellow
    }
    
    # ============================================
    # 3. DEVICE CONFIGURATION PROFILES
    # ============================================
    Write-Host "[3/7] Checking Device Configuration Profiles..." -ForegroundColor Cyan
    
    try {
        $ConfigPolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations" -Method GET -ErrorAction Stop
        foreach ($Policy in $ConfigPolicies.value) {
            $Details = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($Policy.id)" -Method GET -ErrorAction Stop
            
            $PINLength = $null
            if ($Details.passwordMinimumLength) { $PINLength = $Details.passwordMinimumLength }
            if ($Details.passcodeMinimumLength) { $PINLength = $Details.passcodeMinimumLength }
            
            if ($PINLength) {
                $AllResults += [PSCustomObject]@{
                    Source              = "Device Configuration Profile"
                    PolicyName          = $Policy.displayName
                    Platform            = $Policy.'@odata.type' -replace '#microsoft.graph.',''
                    PINRequired         = $Details.passwordRequired
                    MinimumPINLength    = $PINLength
                    PINCharacterSet     = $Details.passwordRequiredType
                    SimplePINBlocked    = $Details.simplePasswordsBlocked
                    PolicyId            = $Policy.id
                    PolicyType          = "ConfigurationProfile"
                }
            }
        }
        Write-Host "  ✓ Found $($ConfigPolicies.value.Count) configuration profiles" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # ============================================
    # 4. SETTINGS CATALOG POLICIES
    # ============================================
    Write-Host "[4/7] Checking Settings Catalog Policies..." -ForegroundColor Cyan
    
    try {
        $SettingsCatalog = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Method GET -ErrorAction Stop
        $PasswordPolicies = 0
        foreach ($Policy in $SettingsCatalog.value) {
            try {
                $Settings = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($Policy.id)/settings" -Method GET -ErrorAction SilentlyContinue
                
                $HasPINSetting = $false
                foreach ($Setting in $Settings.value) {
                    if ($Setting.settingDefinitionId -match "pin|password|passcode") {
                        $HasPINSetting = $true
                        break
                    }
                }
                
                if ($HasPINSetting) {
                    $PasswordPolicies++
                    $AllResults += [PSCustomObject]@{
                        Source              = "Settings Catalog"
                        PolicyName          = $Policy.name
                        Platform            = $Policy.platforms -join ","
                        PINRequired         = "Check in Intune Portal"
                        MinimumPINLength    = "Check in Intune Portal"
                        PINCharacterSet     = "Check in Intune Portal"
                        SimplePINBlocked    = "Check in Intune Portal"
                        PolicyId            = $Policy.id
                        PolicyType          = "SettingsCatalog"
                    }
                }
            }
            catch {
                # Skip policies we can't read
                continue
            }
        }
        Write-Host "  ✓ Found $($SettingsCatalog.value.Count) total, $PasswordPolicies with PIN settings" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # ============================================
    # 5. CONDITIONAL ACCESS POLICIES
    # ============================================
    Write-Host "[5/7] Checking Conditional Access Policies..." -ForegroundColor Cyan
    
    try {
        $CAPolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Method GET -ErrorAction Stop
        $CACount = 0
        foreach ($Policy in $CAPolicies.value) {
            if ($Policy.grantControls.builtInControls -contains "approvedApplication" -or
                $Policy.grantControls.builtInControls -contains "compliantApplication" -or
                $Policy.grantControls.builtInControls -contains "compliantDevice") {
                $CACount++
                $AllResults += [PSCustomObject]@{
                    Source              = "Conditional Access"
                    PolicyName          = $Policy.displayName
                    Platform            = "All"
                    PINRequired         = "Enforces App Protection/Compliance"
                    MinimumPINLength    = "Inherited from other policies"
                    PINCharacterSet     = "N/A"
                    SimplePINBlocked    = "N/A"
                    PolicyId            = $Policy.id
                    PolicyType          = "ConditionalAccess"
                }
            }
        }
        Write-Host "  ✓ Found $CACount CA policies enforcing app protection" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Required permission: Policy.Read.All" -ForegroundColor Yellow
    }
    
    # ============================================
    # 6. APP CONFIGURATION POLICIES
    # ============================================
    Write-Host "[6/7] Checking App Configuration Policies..." -ForegroundColor Cyan
    
    try {
        $AppConfigPolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations" -Method GET -ErrorAction Stop
        Write-Host "  ✓ Found $($AppConfigPolicies.value.Count) app configuration policies" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # ============================================
    # 7. ENROLLMENT RESTRICTIONS
    # ============================================
    Write-Host "[7/7] Checking Enrollment Restrictions..." -ForegroundColor Cyan
    
    try {
        $EnrollmentRestrictions = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations" -Method GET -ErrorAction Stop
        foreach ($Restriction in $EnrollmentRestrictions.value) {
            if ($Restriction.passwordMinimumLength) {
                $AllResults += [PSCustomObject]@{
                    Source              = "Enrollment Restriction"
                    PolicyName          = $Restriction.displayName
                    Platform            = $Restriction.'@odata.type' -replace '#microsoft.graph.',''
                    PINRequired         = $Restriction.passwordRequired
                    MinimumPINLength    = $Restriction.passwordMinimumLength
                    PINCharacterSet     = $Restriction.passwordRequiredType
                    SimplePINBlocked    = $Restriction.simplePasswordBlocked
                    PolicyId            = $Restriction.id
                    PolicyType          = "EnrollmentRestriction"
                }
            }
        }
        Write-Host "  ✓ Found $($EnrollmentRestrictions.value.Count) enrollment restrictions" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # ============================================
    # DISPLAY RESULTS
    # ============================================
    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host " RESULTS" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    
    if ($AllResults.Count -eq 0) {
        Write-Host "⚠ No PIN/Password policies found (or insufficient permissions)" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Required Azure AD roles:" -ForegroundColor Cyan
        Write-Host "  - Global Administrator" -ForegroundColor White
        Write-Host "  - Intune Administrator" -ForegroundColor White
        Write-Host "  - Cloud Device Administrator (read-only)" -ForegroundColor White
    } else {
        Write-Host "Total policies found: $($AllResults.Count)" -ForegroundColor Green
        Write-Host ""
        
        # Highlight 8+ digit requirements
        $EightDigitPolicies = $AllResults | Where-Object { 
            $_.MinimumPINLength -match '^\d+$' -and [int]$_.MinimumPINLength -ge 8
        }
        
        if ($EightDigitPolicies) {
            Write-Host "⚠⚠⚠ POLICIES ENFORCING 8+ DIGIT PIN ⚠⚠⚠" -ForegroundColor Red
            Write-Host ""
            $EightDigitPolicies | Format-Table -AutoSize
            Write-Host ""
        }
        
        # Group by source
        Write-Host "All Policies by Source:" -ForegroundColor Cyan
        $AllResults | Sort-Object Source, PolicyName | Format-Table -AutoSize
        
        # Export
        Write-Host ""
        $ExportChoice = Read-Host "Export full results to CSV? (Y/N)"
        if ($ExportChoice -match "^[Yy]") {
            if (-not (Test-Path "C:\Temp")) {
                New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null
            }
            $ExportPath = "C:\Temp\TenantPINAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
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
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    exit 1
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Host ""
    Write-Host "Disconnected from Microsoft Graph." -ForegroundColor DarkGray
}

<# Usage:
.\PullPINSettings.ps1
.\PullPINSettings.ps1 -AdminUPN "admin@company.com"

Required Azure AD Roles:
- Global Administrator (full access)
- Intune Administrator (full access)
- Cloud Device Administrator (read-only)
#>