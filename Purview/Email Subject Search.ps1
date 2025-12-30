param(
    [string]$Mailbox = "paula.tang@smaileskrawitz.com.au",
    [string]$Sender = "",
    [string]$AttachmentName = "Instruction Sheet – Discretionary Trust Deed Horst Maberly",
    [datetime]$FromDate = [datetime]'2022-01-01',
    [datetime]$ToDate = [datetime]'2022-12-31',
    [switch]$Export,
    [Parameter(Mandatory = $true)]
    [string]$AdminUPN
)

$ErrorActionPreference = 'Stop'

function Reset-PurviewSession {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    try {
        $sessions = Get-PSSession | Where-Object {
            $_.ConfigurationName -eq 'Microsoft.Exchange' -and (
                $_.ComputerName -match 'ps\.compliance\.protection\.outlook' -or
                $_.ComputerName -match 'outlook\.office365'
            )
        }
        if ($sessions) { Remove-PSSession -Session $sessions -ErrorAction SilentlyContinue }
    } catch {}
}

function Ensure-Module {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser -AllowClobber
    }
    Import-Module ExchangeOnlineManagement -Force
}

function Connect-Purview {
    param([string]$UPN)
    Reset-PurviewSession
    Ensure-Module
    Write-Host "Connecting to Purview as $UPN..." -ForegroundColor Yellow
    $cmdParams = @{ UserPrincipalName = $UPN }
    try {
        Connect-IPPSSession @cmdParams
        if (Get-Command New-ComplianceSearch -ErrorAction SilentlyContinue) { return }
    } catch { Write-Host "Default connect failed, retrying..." -ForegroundColor DarkYellow }
    if ((Get-Command Connect-IPPSSession).Parameters.ContainsKey('UseRPSSession')) {
        try {
            Reset-PurviewSession
            Connect-IPPSSession @cmdParams -UseRPSSession
            if (Get-Command New-ComplianceSearch -ErrorAction SilentlyContinue) { return }
        } catch { Write-Host "RPS connect failed..." -ForegroundColor DarkYellow }
    }
    throw "Could not load Compliance cmdlets."
}

Connect-Purview -UPN $AdminUPN

$receivedFrom = $FromDate.ToString('yyyy-MM-dd')
$receivedTo   = $ToDate.ToString('yyyy-MM-dd')
$timestamp    = (Get-Date -Format 'yyyyMMdd_HHmmss')

# Build KQL
$queryParts = @(
    "kind:email",
    "received>=$receivedFrom",
    "received<=$receivedTo"
)
if ($AttachmentName) {
    $escAttach = $AttachmentName.Replace('"','\"')
    $queryParts += @("hasattachment:true","attachment:`"$escAttach`"","attachment:.pdf")
}
$query = ($queryParts -join " AND ")

$slug = "Paula_2022_Horst_Maberly"
$SearchName = "EmailSearch_${slug}_$timestamp"

Write-Host "Search: $Mailbox | $receivedFrom to $receivedTo" -ForegroundColor Green
Write-Host "KQL: $query" -ForegroundColor DarkCyan

function Invoke-ContentSearch {
    param([string]$Name, [string]$Location, [string]$Query, [switch]$Export)
    $existing = Get-ComplianceSearch -Identity $Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "Removing existing search '$Name'..." -ForegroundColor Yellow
        Remove-ComplianceSearch -Identity $Name -Confirm:$false
        Start-Sleep -Seconds 2
    }
    try {
        New-ComplianceSearch -Name $Name -ExchangeLocation $Location -ContentMatchQuery $Query | Out-Null
    } catch {
        Write-Host "New-ComplianceSearch failed. Check mailbox: $Location" -ForegroundColor Red
        throw
    }
    Start-ComplianceSearch -Identity $Name | Out-Null
    Write-Host "Started. Waiting..." -ForegroundColor Yellow
    do {
        Start-Sleep -Seconds 10
        $s = Get-ComplianceSearch -Identity $Name
        Write-Host "Status: $($s.Status) | Items: $($s.Items)" -ForegroundColor Cyan
    } while ($s.Status -ne "Completed")
    if ($Export) {
        $hasExport = (Get-Command New-ComplianceSearchAction).Parameters.Keys -contains 'Export'
        if ($hasExport) {
            New-ComplianceSearchAction -SearchName $Name -Export -Format FxStream | Out-Null
            Write-Host "Export queued. Download from Purview portal." -ForegroundColor Green
        } else {
            Write-Host "Export unavailable. Use portal." -ForegroundColor Yellow
        }
    } else {
        Write-Host "Search completed (no export)." -ForegroundColor Green
    }
}

try {
    Invoke-ContentSearch -Name $SearchName -Location $Mailbox -Query $query -Export:$Export
} finally {
    Reset-PurviewSession
    Write-Host "Session cleaned." -ForegroundColor DarkGray
}

Write-Host "Done." -ForegroundColor Green

<# Usage:
.\Email Subject Search.ps1 -AdminUPN "admin@domain.com" -Export
#>