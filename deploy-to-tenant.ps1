#Requires -Version 5.1
<#
.SYNOPSIS
    Deploys the Outlook Agenda Enforcer manifest.xml to the xtwjy.onmicrosoft.com dev tenant
    via Exchange Online Centralized Deployment.

.NOTES
    Requires: ExchangeOnlineManagement PowerShell module (installed automatically if missing)
    Run as: Global Admin or Exchange Admin on xtwjy.onmicrosoft.com
#>

$ManifestPath = "$PSScriptRoot\manifest.xml"
$TenantDomain = "xtwjy.onmicrosoft.com"
$AddInId      = "33e08f87-3205-4672-9721-c49b6b9dc6f8"

# Any orphaned IDs left behind by earlier failed upload attempts
$OrphanIds    = @("06d8510b-002e-407a-bab9-1894cc825687")

# ── 1. Ensure ExchangeOnlineManagement is available ──────────────────────────
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Cyan
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# ── 2. Connect to Exchange Online ────────────────────────────────────────────
Write-Host "Connecting to Exchange Online ($TenantDomain) - sign in with your admin account when prompted..." -ForegroundColor Cyan
Connect-ExchangeOnline -Organization $TenantDomain -ShowBanner:$false

# ── 3. Read manifest bytes ────────────────────────────────────────────────────
if (-not (Test-Path $ManifestPath)) {
    Write-Error "manifest.xml not found at: $ManifestPath"
    exit 1
}
$manifestBytes = [System.IO.File]::ReadAllBytes($ManifestPath)
Write-Host "Manifest loaded: $ManifestPath" -ForegroundColor Green

# ── 4. Show everything currently deployed (diagnostic) ───────────────────────
Write-Host "`nCurrently deployed org add-ins:" -ForegroundColor Cyan
$allApps = Get-App -OrganizationApp -ErrorAction SilentlyContinue
if ($allApps) {
    $allApps | Select-Object DisplayName, AppId, Enabled | Format-Table -AutoSize
} else {
    Write-Host "  (none found)" -ForegroundColor Gray
}

# ── 5. Remove any orphaned records from previous failed upload attempts ────────
foreach ($orphanId in $OrphanIds) {
    $orphan = $allApps | Where-Object { $_.AppId -eq $orphanId }
    if ($orphan) {
        Write-Host "Removing orphaned add-in (AppId: $orphanId)..." -ForegroundColor Yellow
        Remove-App -OrganizationApp -Identity $orphanId -Confirm:$false
        Start-Sleep -Seconds 3
        Write-Host "Orphan removed." -ForegroundColor Green
    }
}

# ── 6. Deploy or update the current add-in ────────────────────────────────────
$existing = Get-App -OrganizationApp -ErrorAction SilentlyContinue | Where-Object { $_.AppId -eq $AddInId }

if ($existing) {
    Write-Host "Add-in already exists (AppId: $AddInId). Updating manifest..." -ForegroundColor Yellow
    # Remove and re-add since Set-App does not support -FileData for manifest replacement
    Remove-App -OrganizationApp -Identity $AddInId -Confirm:$false
    Write-Host "Waiting for removal to propagate..." -ForegroundColor Yellow
    Start-Sleep -Seconds 5
    New-App -OrganizationApp `
        -FileData $manifestBytes `
        -DefaultStateForUser Enabled `
        -ProvidedTo Everyone `
        -ErrorAction Stop
    Write-Host "Add-in updated successfully." -ForegroundColor Green
} else {
    Write-Host "Deploying new add-in to tenant..." -ForegroundColor Cyan
    New-App -OrganizationApp `
        -FileData $manifestBytes `
        -DefaultStateForUser Enabled `
        -ProvidedTo Everyone `
        -ErrorAction Stop
    Write-Host "Add-in deployed successfully." -ForegroundColor Green
}

# ── 7. Verify deployment ─────────────────────────────────────────────────────
Write-Host "`nVerifying deployment..." -ForegroundColor Cyan
$deployed = Get-App -OrganizationApp -ErrorAction SilentlyContinue | Where-Object { $_.AppId -eq $AddInId }

if ($deployed) {
    Write-Host "`n✔  Deployment confirmed:" -ForegroundColor Green
    $deployed | Select-Object DisplayName, AppId, Enabled, DefaultStateForUser, ProvidedTo | Format-List
} else {
    Write-Warning "Could not verify deployment. Check the Microsoft 365 Admin Center under Settings > Integrated apps."
}

# ── 8. Disconnect ────────────────────────────────────────────────────────────
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Done. Remember to start your local dev server (npm start) before testing in Outlook." -ForegroundColor Cyan
