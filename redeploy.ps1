# redeploy.ps1
# Removes and re-registers the Outlook Agenda Enforcer add-in in Exchange Online.
# Run this from the repo root after uploading dist/commands-v10.html and dist/launchevent-v10.js to GitHub Pages.

param(
    [Parameter(Mandatory=$true)]
    [string]$AdminUpn   # e.g. admin@xtwjy.onmicrosoft.com
)

$ManifestPath = Join-Path $PSScriptRoot "dist\manifest.xml"

if (-not (Test-Path $ManifestPath)) {
    Write-Error "dist\manifest.xml not found. Run 'npm run build' first."
    exit 1
}

# Verify ExchangeOnlineManagement module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
}

Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName $AdminUpn

Write-Host "`n=== Current add-in registrations ===" -ForegroundColor Cyan
$existing = Get-OrganizationAddIn | Where-Object { $_.DisplayName -like "*Agenda*" -or $_.AppId -eq "33e08f87-3205-4672-9721-c49b6b9dc6f8" }
$existing | Format-Table DisplayName, AppId, Version, Enabled -AutoSize

if ($existing) {
    foreach ($addIn in $existing) {
        Write-Host "Removing: $($addIn.DisplayName) v$($addIn.Version) ($($addIn.AppId))" -ForegroundColor Yellow
        Remove-OrganizationAddIn -Identity $addIn.AppId -Confirm:$false
        Write-Host "Removed." -ForegroundColor Green
    }
} else {
    Write-Host "No existing Agenda Enforcer registration found." -ForegroundColor Yellow
}

Write-Host "`nReading manifest: $ManifestPath" -ForegroundColor Cyan
$manifestBytes = [System.IO.File]::ReadAllBytes($ManifestPath)

# Confirm version in manifest
$manifestXml = [xml][System.Text.Encoding]::UTF8.GetString($manifestBytes)
$version = $manifestXml.OfficeApp.Version
Write-Host "Manifest version: $version" -ForegroundColor Cyan

Write-Host "Registering add-in..." -ForegroundColor Cyan
$newAddIn = New-OrganizationAddIn -FileData $manifestBytes -DefaultStateForUser Enabled
Write-Host "Registered: $($newAddIn.DisplayName) v$($newAddIn.Version)" -ForegroundColor Green

Write-Host "`n=== Updated registrations ===" -ForegroundColor Cyan
Get-OrganizationAddIn | Where-Object { $_.DisplayName -like "*Agenda*" -or $_.AppId -eq "33e08f87-3205-4672-9721-c49b6b9dc6f8" } | Format-Table DisplayName, AppId, Version, Enabled -AutoSize

Disconnect-ExchangeOnline -Confirm:$false
Write-Host "`nDone. Now clear ALL browser data in OWA and reload." -ForegroundColor Green
