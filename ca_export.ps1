<#
.SYNOPSIS
Exports all Microsoft Entra Conditional Access (CA) policies to a readable CSV overview
and/or a full-fidelity JSON file, resolving GUIDs to display names.

.DESCRIPTION
Connects to Microsoft Graph and retrieves every Conditional Access policy. GUID
references (users, groups, directory roles, applications, named locations) are resolved
to display names on demand, so the script does not need to download the entire
directory. Output format is selectable:

  CSV   A flattened, human-readable overview, one row per policy.
  JSON  The raw Graph representation of each policy, suitable for backup, diffing and
        re-import.
  Both  Writes both files (default).

Only the Microsoft.Graph.Authentication module is required; everything runs through
Invoke-MgGraphRequest.

.PARAMETER OutputFormat
CSV, JSON, or Both. Default: Both.

.PARAMETER OutputDirectory
Folder to write the output files to. Default: the current directory.

.PARAMETER CsvDelimiter
Delimiter for the CSV file. Default: ",". Use ";" for locales where Excel expects a
semicolon (for example Norwegian).

.PARAMETER UseDeviceCode
Sign in with device code flow instead of the interactive browser. Useful if interactive
sign-in fails with a broker/window-handle error, or on a headless host.

.NOTES
Requirements:
  - PowerShell 7 or later.
  - The Microsoft.Graph.Authentication module:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  - Delegated scopes Policy.Read.All and Directory.Read.All (consented on first run).

.EXAMPLE
.\ca-export.ps1

.EXAMPLE
.\ca-export.ps1 -OutputFormat JSON

.EXAMPLE
.\ca-export.ps1 -OutputFormat CSV -CsvDelimiter ";" -OutputDirectory C:\exports
#>

[CmdletBinding()]
param(
    [ValidateSet('CSV','JSON','Both')]
    [string]$OutputFormat = 'Both',

    [string]$OutputDirectory = (Get-Location).Path,

    [string]$CsvDelimiter = ',',

    [switch]$UseDeviceCode
)

#--------------------------------------------------------------------
# Logging helpers
#--------------------------------------------------------------------
function Write-Step {
    param([string]$Message)
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host "[$ts] $Message"
}

function Measure-Step {
    param([string]$Name, [scriptblock]$Block)
    Write-Step "$Name (start)"
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $result = & $Block
    $sw.Stop()
    Write-Step "$Name (done) - $([math]::Round($sw.Elapsed.TotalSeconds,1))s"
    return $result
}

#--------------------------------------------------------------------
# Graph helpers
#--------------------------------------------------------------------
function Get-GraphValue {
    # Single GET, returns $null on failure so name resolution degrades gracefully.
    # Retries a few times on throttling (429).
    param([string]$Uri, [int]$MaxRetries = 3)
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
        }
        catch {
            $code = 0
            try { $code = [int]$_.Exception.Response.StatusCode } catch {}
            if ($code -eq 429 -and $attempt -le $MaxRetries) {
                Start-Sleep -Seconds ([Math]::Min(20, [int][Math]::Pow(2, $attempt)))
                continue
            }
            return $null
        }
    }
}

function Get-AllPages {
    # GET with @odata.nextLink paging.
    param([string]$Uri)
    $items = @()
    $next = $Uri
    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject -ErrorAction Stop
        if ($resp.value) { $items += $resp.value }
        $next = $resp.'@odata.nextLink'
    } while ($next)
    return $items
}

#--------------------------------------------------------------------
# Formatting helpers
#--------------------------------------------------------------------
$MultiSep = ' | '

function Format-Multi {
    param([object[]]$Values)
    if (-not $Values) { return "" }
    ($Values | Where-Object { $null -ne $_ -and "$_" -ne "" } | ForEach-Object { "$_" }) -join $MultiSep
}

function Format-DeviceFilter {
    param($Devices)
    if (-not $Devices -or -not $Devices.deviceFilter) { return "" }
    "{0}: {1}" -f $Devices.deviceFilter.mode, $Devices.deviceFilter.rule
}

function Format-SessionControls {
    param($Sc)
    if (-not $Sc) { return "" }
    $parts = @()
    if ($Sc.applicationEnforcedRestrictions -and $Sc.applicationEnforcedRestrictions.isEnabled) {
        $parts += "AppEnforcedRestrictions"
    }
    if ($Sc.cloudAppSecurity -and $Sc.cloudAppSecurity.isEnabled) {
        $parts += "CloudAppSecurity=$($Sc.cloudAppSecurity.cloudAppSecurityType)"
    }
    if ($Sc.signInFrequency -and $Sc.signInFrequency.isEnabled) {
        $sif = $Sc.signInFrequency
        $val = if ($sif.frequencyInterval -eq 'everyTime') { 'everyTime' } else { "$($sif.value) $($sif.type)" }
        $parts += "SignInFrequency=$val"
    }
    if ($Sc.persistentBrowser -and $Sc.persistentBrowser.isEnabled) {
        $parts += "PersistentBrowser=$($Sc.persistentBrowser.mode)"
    }
    if ($Sc.continuousAccessEvaluation -and $Sc.continuousAccessEvaluation.mode) {
        $parts += "CAE=$($Sc.continuousAccessEvaluation.mode)"
    }
    if ($null -ne $Sc.disableResilienceDefaults) {
        $parts += "DisableResilienceDefaults=$($Sc.disableResilienceDefaults)"
    }
    $parts -join $MultiSep
}

#--------------------------------------------------------------------
# Resolvers (cached, targeted lookups)
#--------------------------------------------------------------------
$script:UserCache  = @{}
$script:GroupCache = @{}
$script:AppCache   = @{}
$script:RoleMap    = @{}   # filled once from directoryRoleTemplates
$script:LocMap     = @{}   # filled once from namedLocations
$script:TouMap     = @{}   # filled once from terms of use agreements

$UserTokens = @{ 'All' = 'All users'; 'None' = 'None'; 'GuestsOrExternalUsers' = 'Guests or external users' }
$AppTokens  = @{ 'All' = 'All cloud apps'; 'None' = 'None'; 'Office365' = 'Office 365'; 'MicrosoftAdminPortals' = 'Microsoft Admin Portals' }
$LocTokens  = @{ 'All' = 'All locations'; 'AllTrusted' = 'All trusted locations' }

function Resolve-DirObject {
    # Resolves a directory object id to a display name. Falls back to deletedItems so
    # stale references to removed users/groups are labelled instead of shown as a GUID.
    param([string]$Id, [string]$Type)  # $Type: 'users' or 'groups'
    $r = Get-GraphValue ("/v1.0/$Type/$Id" + '?$select=displayName')
    if ($r -and $r.displayName) { return $r.displayName }
    $d = Get-GraphValue ("/v1.0/directory/deletedItems/$Id" + '?$select=displayName')
    if ($d -and $d.displayName) { return "$($d.displayName) (deleted)" }
    return "$Id (not found)"
}

function Resolve-User {
    param([string[]]$Ids)
    if (-not $Ids) { return "" }
    $out = foreach ($id in $Ids) {
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        if ($UserTokens.ContainsKey($id)) { $UserTokens[$id]; continue }
        if ($script:UserCache.ContainsKey($id)) { $script:UserCache[$id]; continue }
        $name = Resolve-DirObject -Id $id -Type 'users'
        $script:UserCache[$id] = $name
        $name
    }
    ($out) -join $MultiSep
}

function Resolve-Group {
    param([string[]]$Ids)
    if (-not $Ids) { return "" }
    $out = foreach ($id in $Ids) {
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        if ($script:GroupCache.ContainsKey($id)) { $script:GroupCache[$id]; continue }
        $name = Resolve-DirObject -Id $id -Type 'groups'
        $script:GroupCache[$id] = $name
        $name
    }
    ($out) -join $MultiSep
}

function Resolve-Role {
    param([string[]]$Ids)
    if (-not $Ids) { return "" }
    $out = foreach ($id in $Ids) {
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        if ($script:RoleMap.ContainsKey($id)) { $script:RoleMap[$id] } else { $id }
    }
    ($out) -join $MultiSep
}

function Resolve-App {
    param([string[]]$Ids)
    if (-not $Ids) { return "" }
    $out = foreach ($id in $Ids) {
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        if ($AppTokens.ContainsKey($id)) { $AppTokens[$id]; continue }
        if ($script:AppCache.ContainsKey($id)) { $script:AppCache[$id]; continue }
        $uri = "/v1.0/servicePrincipals" + '?$filter=' + "appId eq '$id'" + '&$select=appId,displayName'
        $r = Get-GraphValue $uri
        $name = if ($r -and $r.value -and $r.value.Count -gt 0 -and $r.value[0].displayName) { $r.value[0].displayName } else { "$id (app not found)" }
        $script:AppCache[$id] = $name
        $name
    }
    ($out) -join $MultiSep
}

function Resolve-Tou {
    param([string[]]$Ids)
    if (-not $Ids) { return "" }
    $out = foreach ($id in $Ids) {
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        if ($script:TouMap.ContainsKey($id)) { $script:TouMap[$id] } else { "$id (not found)" }
    }
    ($out) -join $MultiSep
}

function Resolve-Location {
    param([string[]]$Ids)
    if (-not $Ids) { return "" }
    $out = foreach ($id in $Ids) {
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        if ($LocTokens.ContainsKey($id)) { $LocTokens[$id]; continue }
        if ($script:LocMap.ContainsKey($id)) { $script:LocMap[$id] } else { $id }
    }
    ($out) -join $MultiSep
}

#====================================================================
# Main
#====================================================================
$globalSw = [System.Diagnostics.Stopwatch]::StartNew()
Write-Step "Starting export"

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Error "Requires PowerShell 7 or later."
    return
}

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Error @"
The Microsoft.Graph.Authentication module is missing. Install it with:
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
and run the script again.
"@
    return
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

if (-not (Test-Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
}

# Connect
$scopes = @('Policy.Read.All','Directory.Read.All','Agreement.Read.All')
Measure-Step "Connect to Microsoft Graph" {
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    $haveAll = $ctx -and -not ($scopes | Where-Object { $ctx.Scopes -notcontains $_ })
    if (-not $haveAll) {
        $cp = @{ Scopes = $scopes; NoWelcome = $true; ErrorAction = 'Stop' }
        if ($UseDeviceCode) { $cp.UseDeviceCode = $true }
        Connect-MgGraph @cp
    }
    $null = Get-MgContext -ErrorAction Stop
}

# Fetch policies
$policies = Measure-Step "Fetch Conditional Access policies" {
    Get-AllPages "/v1.0/identity/conditionalAccess/policies"
}
Write-Step "Policies: $($policies.Count)"

# Prefetch bounded lookup sets (roles, named locations, terms of use)
Measure-Step "Fetch role templates, named locations and terms of use" {
    Get-AllPages ("/v1.0/directoryRoleTemplates" + '?$select=id,displayName') |
        ForEach-Object { if ($_.id) { $script:RoleMap[$_.id] = $_.displayName } }
    Get-AllPages ("/v1.0/identity/conditionalAccess/namedLocations" + '?$select=id,displayName') |
        ForEach-Object { if ($_.id) { $script:LocMap[$_.id] = $_.displayName } }

    # Terms of use: endpoint path/permission varies, so try a couple and fail quietly.
    foreach ($touUri in @('/v1.0/agreements?$select=id,displayName',
                          '/beta/identityGovernance/termsOfUse/agreements?$select=id,displayName')) {
        try {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $touUri -OutputType PSObject -ErrorAction Stop
            if ($resp.value) {
                $resp.value | ForEach-Object { if ($_.id) { $script:TouMap[$_.id] = $_.displayName } }
                break
            }
        } catch { }
    }
}

# Build the readable overview (users, groups and apps are resolved on demand and cached)
$overview = Measure-Step "Build overview rows" {
    foreach ($p in $policies) {
        [pscustomobject][ordered]@{
            Id               = $p.id
            DisplayName      = $p.displayName
            State            = $p.state
            IncludeUsers     = Resolve-User     $p.conditions.users.includeUsers
            ExcludeUsers     = Resolve-User     $p.conditions.users.excludeUsers
            IncludeGroups    = Resolve-Group    $p.conditions.users.includeGroups
            ExcludeGroups    = Resolve-Group    $p.conditions.users.excludeGroups
            IncludeRoles     = Resolve-Role     $p.conditions.users.includeRoles
            ExcludeRoles     = Resolve-Role     $p.conditions.users.excludeRoles
            IncludeApps      = Resolve-App      $p.conditions.applications.includeApplications
            ExcludeApps      = Resolve-App      $p.conditions.applications.excludeApplications
            UserActions      = Format-Multi     $p.conditions.applications.includeUserActions
            AuthContexts     = Format-Multi     $p.conditions.applications.includeAuthenticationContextClassReferences
            ClientAppTypes   = Format-Multi     $p.conditions.clientAppTypes
            IncludePlatforms = Format-Multi     $p.conditions.platforms.includePlatforms
            ExcludePlatforms = Format-Multi     $p.conditions.platforms.excludePlatforms
            IncludeLocations = Resolve-Location $p.conditions.locations.includeLocations
            ExcludeLocations = Resolve-Location $p.conditions.locations.excludeLocations
            DeviceFilter     = Format-DeviceFilter $p.conditions.devices
            SignInRiskLevels = Format-Multi     $p.conditions.signInRiskLevels
            UserRiskLevels   = Format-Multi     $p.conditions.userRiskLevels
            GrantOperator    = $p.grantControls.operator
            GrantControls    = Format-Multi     $p.grantControls.builtInControls
            AuthStrength     = $p.grantControls.authenticationStrength.displayName
            TermsOfUse       = Resolve-Tou      $p.grantControls.termsOfUse
            SessionControls  = Format-SessionControls $p.sessionControls
        }
    }
}

# Write output
$stamp    = Get-Date -Format "yyyyMMdd-HHmmss"
$csvPath  = Join-Path $OutputDirectory "ConditionalAccess-overview_$stamp.csv"
$jsonPath = Join-Path $OutputDirectory "ConditionalAccess-policies_$stamp.json"

if ($OutputFormat -in 'CSV','Both') {
    Measure-Step "Write CSV" {
        $overview | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter $CsvDelimiter
    }
    Write-Step "CSV: $csvPath"
}

if ($OutputFormat -in 'JSON','Both') {
    Measure-Step "Write JSON" {
        $policies | ConvertTo-Json -Depth 25 | Out-File -FilePath $jsonPath -Encoding UTF8
    }
    Write-Step "JSON: $jsonPath"
}

$globalSw.Stop()
Write-Step "Finished - $($policies.Count) policies in $([math]::Round($globalSw.Elapsed.TotalSeconds,1))s"
