<#
.SYNOPSIS
    Advanced Ethernet disconnect analysis with detailed event correlation
.DESCRIPTION
    Distinguishes between real disconnects and virtual adapter state changes
#>

param(
    [Parameter(Mandatory = $false)]
    [switch]$ContinuousMonitor,
    
    [Parameter(Mandatory = $false)]
    [int]$MonitorDurationMinutes = 60,
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowAllEvents
)

$ErrorActionPreference = 'Continue'

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " ADVANCED ETHERNET DISCONNECT ANALYZER" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

$LogPath = "C:\Temp\NetworkDiagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

if (-not (Test-Path "C:\Temp")) {
    New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null
}

Start-Transcript -Path $LogPath -Append

# ============================================
# IDENTIFY REAL ETHERNET ADAPTERS
# ============================================
Write-Host "[1/5] Identifying Physical Ethernet Adapters..." -ForegroundColor Cyan

$AllAdapters = Get-NetAdapter | Sort-Object Name
$PhysicalEthernet = @()
$VirtualAdapters = @()

foreach ($Adapter in $AllAdapters) {
    # More precise filtering for physical Ethernet
    $IsVirtual = $false
    
    # Check name patterns
    if ($Adapter.Name -match "vEthernet|Loopback|VPN|TAP|Hyper|VMware|Virtual|Bluetooth|WiFi") {
        $IsVirtual = $true
    }
    
    # Check description patterns  
    if ($Adapter.InterfaceDescription -match "Virtual|Hyper|VMware|Bluetooth|Wireless|WiFi") {
        $IsVirtual = $true
    }
    
    # Check media type
    if ($Adapter.PhysicalMediaType -ne "802.3") {
        $IsVirtual = $true
    }
    
    if (-not $IsVirtual) {
        $PhysicalEthernet += $Adapter
        Write-Host "  + Physical: $($Adapter.Name) - $($Adapter.InterfaceDescription)" -ForegroundColor Green
    } else {
        $VirtualAdapters += $Adapter
        if ($ShowAllEvents) {
            Write-Host "  - Virtual/Other: $($Adapter.Name) - $($Adapter.InterfaceDescription)" -ForegroundColor DarkGray
        }
    }
}

if ($PhysicalEthernet.Count -eq 0) {
    Write-Host "  X No physical Ethernet adapters found!" -ForegroundColor Red
    exit 1
}

Write-Host "  Found $($PhysicalEthernet.Count) physical Ethernet adapter(s)" -ForegroundColor Cyan
Write-Host ""

# ============================================
# ANALYZE EVENT LOGS WITH CORRELATION
# ============================================
Write-Host "[2/5] Analyzing Network Events (Last 24 Hours)..." -ForegroundColor Cyan

$StartDate = (Get-Date).AddDays(-1)
$NetworkEvents = @()

# Get comprehensive network events
$EventIDs = @{
    27 = "Network adapter disabled"
    32 = "Network adapter enabled" 
    33 = "Network adapter reconnected"
    4202 = "Network location changed"
    4226 = "TCP/IP connection limit reached"
    5719 = "Domain controller connection failed"
    1014 = "DNS resolution failed"
    1129 = "Network interface status changed"
}

foreach ($EventID in $EventIDs.Keys) {
    $Events = Get-WinEvent -FilterHashtable @{
        LogName = 'System'
        ID = $EventID
        StartTime = $StartDate
    } -ErrorAction SilentlyContinue
    
    foreach ($Event in $Events) {
        $NetworkEvents += [PSCustomObject]@{
            TimeCreated = $Event.TimeCreated
            EventID = $Event.Id
            Description = $EventIDs[$Event.Id]
            Message = $Event.Message
            Level = $Event.LevelDisplayName
            Source = $Event.ProviderName
        }
    }
}

# Sort by time
$NetworkEvents = $NetworkEvents | Sort-Object TimeCreated -Descending

Write-Host "  Found $($NetworkEvents.Count) network-related events" -ForegroundColor White

# Analyze patterns
$DisconnectEvents = $NetworkEvents | Where-Object { $_.EventID -in @(27, 1129) }
$ReconnectEvents = $NetworkEvents | Where-Object { $_.EventID -in @(32, 33) }

if ($DisconnectEvents) {
    Write-Host ""
    Write-Host "  DISCONNECT ANALYSIS:" -ForegroundColor Yellow
    
    foreach ($Event in ($DisconnectEvents | Select-Object -First 10)) {
        # Try to identify which adapter
        $AdapterName = "Unknown"
        
        # Simple regex patterns to avoid complex parsing
        if ($Event.Message -like "*adapter*") {
            $MessageParts = $Event.Message -split "'"
            if ($MessageParts.Count -gt 1) {
                $AdapterName = $MessageParts[1]
            }
        }
        
        # Check if it's a physical adapter
        $IsPhysical = $PhysicalEthernet | Where-Object { 
            $_.Name -eq $AdapterName -or $_.InterfaceDescription -like "*$AdapterName*" 
        }
        
        $Color = if ($IsPhysical) { "Red" } else { "DarkGray" }
        $Type = if ($IsPhysical) { "[PHYSICAL]" } else { "[VIRTUAL]" }
        
        Write-Host "    $($Event.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')) $Type $AdapterName" -ForegroundColor $Color
        
        if ($IsPhysical) {
            # Look for correlated events within 30 seconds
            $CorrelatedEvents = $NetworkEvents | Where-Object { 
                $_.TimeCreated -gt $Event.TimeCreated.AddSeconds(-30) -and 
                $_.TimeCreated -lt $Event.TimeCreated.AddSeconds(30) -and
                $_.EventID -ne $Event.EventID
            }
            
            if ($CorrelatedEvents) {
                Write-Host "      Correlated events:" -ForegroundColor Yellow
                foreach ($Corr in $CorrelatedEvents) {
                    Write-Host "        $($Corr.TimeCreated.ToString('HH:mm:ss')) - $($Corr.Description)" -ForegroundColor DarkYellow
                }
            }
        }
    }
}

Write-Host ""

# ============================================
# CHECK POWER MANAGEMENT (DETAILED)
# ============================================
Write-Host "[3/5] Detailed Power Management Analysis..." -ForegroundColor Cyan

foreach ($Adapter in $PhysicalEthernet) {
    Write-Host ""
    Write-Host "  $($Adapter.Name):" -ForegroundColor Yellow
    
    # Get device instance path
    $PnPDevice = Get-PnpDevice | Where-Object { 
        $_.FriendlyName -eq $Adapter.InterfaceDescription -or 
        $_.Name -eq $Adapter.InterfaceDescription 
    } | Select-Object -First 1
    
    if ($PnPDevice) {
        Write-Host "    Device Instance: $($PnPDevice.InstanceId)" -ForegroundColor DarkGray
        
        # Check power capabilities
        try {
            $PowerCaps = powercfg /devicequery wake_armed 2>$null | Where-Object { $_ -match [regex]::Escape($PnPDevice.InstanceId) }
            if ($PowerCaps) {
                Write-Host "    ! Device can wake system (may indicate power management issues)" -ForegroundColor Red
            } else {
                Write-Host "    + Device not configured for wake" -ForegroundColor Green
            }
        } catch {
            Write-Host "    Could not check wake capabilities" -ForegroundColor DarkGray
        }
    }
    
    # Check advanced properties that can cause disconnects
    $AdvancedProps = @{
        "*EEE" = "Energy Efficient Ethernet"
        "*FlowControl" = "Flow Control" 
        "*InterruptModeration" = "Interrupt Moderation"
        "*RSS" = "Receive Side Scaling"
        "*TCPConnectionOffload" = "TCP Connection Offload"
        "*USOv2IPv4" = "Large Send Offload v2 (IPv4)"
        "*USOv2IPv6" = "Large Send Offload v2 (IPv6)"
    }
    
    foreach ($Prop in $AdvancedProps.Keys) {
        try {
            $Setting = Get-NetAdapterAdvancedProperty -Name $Adapter.Name -RegistryKeyword $Prop -ErrorAction SilentlyContinue
            if ($Setting) {
                $Status = switch ($Setting.RegistryValue) {
                    0 { "Disabled" }
                    1 { "Enabled" }
                    default { $Setting.DisplayValue }
                }
                
                $Color = switch ($Prop) {
                    "*EEE" { if ($Setting.RegistryValue -eq 1) { "Yellow" } else { "Green" } }
                    default { "White" }
                }
                
                Write-Host "    $($AdvancedProps[$Prop]): $Status" -ForegroundColor $Color
                
                if ($Prop -eq "*EEE" -and $Setting.RegistryValue -eq 1) {
                    Write-Host "      ! EEE can cause intermittent disconnects" -ForegroundColor Yellow
                }
            }
        } catch {
            # Property not available for this adapter
        }
    }
}

Write-Host ""

# ============================================
# ENHANCED CONTINUOUS MONITORING
# ============================================
if ($ContinuousMonitor) {
    Write-Host "[4/5] Starting Enhanced Continuous Monitoring..." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Monitoring for $MonitorDurationMinutes minutes..." -ForegroundColor Yellow
    Write-Host "This will detect REAL disconnects vs virtual adapter changes" -ForegroundColor White
    Write-Host "Press Ctrl+C to stop early" -ForegroundColor DarkGray
    Write-Host ""
    
    $MonitorLog = "C:\Temp\NetworkMonitor_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $EventLog = "C:\Temp\NetworkEvents_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    # Initialize logs
    "Timestamp,Adapter,Status,LinkSpeed,MediaConnectState,OperationalStatus,AdminStatus,LastChange" | Out-File $MonitorLog
    "Timestamp,EventType,Adapter,Details,Severity" | Out-File $EventLog
    
    $StartTime = Get-Date
    $EndTime = $StartTime.AddMinutes($MonitorDurationMinutes)
    $LastStates = @{}
    
    # Baseline current states
    foreach ($Adapter in $PhysicalEthernet) {
        $Current = Get-NetAdapter -Name $Adapter.Name
        $LastStates[$Adapter.Name] = @{
            Status = $Current.Status
            MediaConnectState = $Current.MediaConnectState
            OperationalStatus = $Current.OperationalStatus
            AdminStatus = $Current.AdminStatus
        }
    }
    
    $CheckCount = 0
    
    while ((Get-Date) -lt $EndTime) {
        $CheckCount++
        $CurrentTime = Get-Date
        
        foreach ($Adapter in $PhysicalEthernet) {
            try {
                $Current = Get-NetAdapter -Name $Adapter.Name
                $Stats = Get-NetAdapterStatistics -Name $Adapter.Name
                
                # Log current state
                $LogEntry = "$($CurrentTime.ToString('yyyy-MM-dd HH:mm:ss')),$($Current.Name),$($Current.Status),$($Current.LinkSpeed),$($Current.MediaConnectState),$($Current.OperationalStatus),$($Current.AdminStatus),$($Current.LastChange)"
                $LogEntry | Out-File $MonitorLog -Append
                
                # Check for state changes
                $LastState = $LastStates[$Adapter.Name]
                $Changes = @()
                
                if ($Current.Status -ne $LastState.Status) {
                    $Changes += "Status: $($LastState.Status) -> $($Current.Status)"
                }
                if ($Current.MediaConnectState -ne $LastState.MediaConnectState) {
                    $Changes += "Media: $($LastState.MediaConnectState) -> $($Current.MediaConnectState)"
                }
                if ($Current.OperationalStatus -ne $LastState.OperationalStatus) {
                    $Changes += "Operational: $($LastState.OperationalStatus) -> $($Current.OperationalStatus)"
                }
                
                if ($Changes.Count -gt 0) {
                    $Severity = "INFO"
                    $Color = "White"
                    
                    # Determine if this is a real disconnect
                    if ($Current.MediaConnectState -eq "Disconnected" -and $LastState.MediaConnectState -eq "Connected") {
                        $Severity = "CRITICAL"
                        $Color = "Red"
                        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] X PHYSICAL DISCONNECT: $($Current.Name)" -ForegroundColor Red
                        
                        # Immediate follow-up checks
                        Start-Sleep -Milliseconds 500
                        $Recheck = Get-NetAdapter -Name $Adapter.Name
                        if ($Recheck.MediaConnectState -eq "Connected") {
                            Write-Host "    -> Reconnected after 500ms (likely cable/port issue)" -ForegroundColor Yellow
                        }
                        
                    } elseif ($Current.Status -eq "Disconnected" -and $Current.MediaConnectState -eq "Connected") {
                        $Severity = "WARNING" 
                        $Color = "Yellow"
                        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] ! SOFTWARE DISCONNECT: $($Current.Name) (media still connected)" -ForegroundColor Yellow
                        
                    } elseif ($Current.Status -eq "Up" -and $LastState.Status -ne "Up") {
                        $Severity = "INFO"
                        $Color = "Green"
                        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] + RECONNECTED: $($Current.Name)" -ForegroundColor Green
                    } else {
                        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] i State Change: $($Current.Name) - $($Changes -join ', ')" -ForegroundColor $Color
                    }
                    
                    # Log the event
                    $EventEntry = "$($CurrentTime.ToString('yyyy-MM-dd HH:mm:ss')),StateChange,$($Current.Name),$($Changes -join '; '),$Severity"
                    $EventEntry | Out-File $EventLog -Append
                    
                    # Update last known state
                    $LastStates[$Adapter.Name] = @{
                        Status = $Current.Status
                        MediaConnectState = $Current.MediaConnectState
                        OperationalStatus = $Current.OperationalStatus
                        AdminStatus = $Current.AdminStatus
                    }
                }
                
            } catch {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] X Error checking $($Adapter.Name): $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        # Show progress every 20 checks (100 seconds)
        if ($CheckCount % 20 -eq 0) {
            $Elapsed = (Get-Date) - $StartTime
            $Remaining = $EndTime - (Get-Date)
            $ElapsedStr = "{0:mm\:ss}" -f $Elapsed
            $RemainingStr = "{0:mm\:ss}" -f $Remaining
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Progress: $ElapsedStr elapsed, $RemainingStr remaining" -ForegroundColor DarkCyan
        }
        
        Start-Sleep -Seconds 5
    }
    
    Write-Host ""
    Write-Host "+ Monitoring complete!" -ForegroundColor Green
    Write-Host "  State log: $MonitorLog" -ForegroundColor White
    Write-Host "  Event log: $EventLog" -ForegroundColor White
} else {
    Write-Host "[4/5] Skipping continuous monitoring (use -ContinuousMonitor to enable)" -ForegroundColor DarkGray
}

# ============================================
# FINAL RECOMMENDATIONS
# ============================================
Write-Host ""
Write-Host "[5/5] Final Analysis & Recommendations..." -ForegroundColor Cyan

$RealDisconnects = $DisconnectEvents | Where-Object { 
    $AdapterName = "Unknown"
    if ($_.Message -like "*adapter*") {
        $MessageParts = $_.Message -split "'"
        if ($MessageParts.Count -gt 1) {
            $AdapterName = $MessageParts[1]
        }
    }
    $PhysicalEthernet | Where-Object { $_.Name -eq $AdapterName -or $_.InterfaceDescription -like "*$AdapterName*" }
}

if ($RealDisconnects.Count -gt 0) {
    Write-Host ""
    Write-Host "! REAL PHYSICAL DISCONNECTS DETECTED: $($RealDisconnects.Count)" -ForegroundColor Red
    Write-Host ""
    Write-Host "RECOMMENDED ACTIONS:" -ForegroundColor Yellow
    Write-Host "  1. Check/replace Ethernet cable" -ForegroundColor White
    Write-Host "  2. Try different switch port" -ForegroundColor White
    Write-Host "  3. Disable Energy Efficient Ethernet (EEE)" -ForegroundColor White
    Write-Host "  4. Update network adapter driver" -ForegroundColor White
    Write-Host "  5. Disable power management on adapter" -ForegroundColor White
} else {
    Write-Host ""
    Write-Host "+ NO REAL PHYSICAL DISCONNECTS FOUND" -ForegroundColor Green
    Write-Host "  Events were likely virtual adapter changes or false positives" -ForegroundColor White
}

Stop-Transcript

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " ANALYSIS COMPLETE" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Full log: $LogPath" -ForegroundColor Green