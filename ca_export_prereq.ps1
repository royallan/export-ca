# ==========================================================
# CA_EXPORT_PREREQUISITES.ps1  (KJØRES ÉN GANG PER MASKIN)
# ==========================================================
# Forutsetter PowerShell 7+ (pwsh).
# Installer Microsoft Graph PowerShell SDK.
# (Valgfritt) Oppdater modulen ved behov.
# ==========================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
  Install-Module Microsoft.Graph -Scope CurrentUser -Force
} else {
  Write-Host "Microsoft.Graph is already installed."
}

# Optional:
# Update-Module Microsoft.Graph

Write-Host "Prerequisites complete."