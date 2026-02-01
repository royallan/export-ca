# ==========================================================
# export-ca-readable.ps1
# Produksjonsklar med progresjonslogging + tidsbruk
# ==========================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Step {
  param(
    [string]$Message
  )
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  Write-Host "[$ts] $Message"
}

function Measure-Step {
  param(
    [string]$Name,
    [scriptblock]$Block
  )
  Write-Step "$Name (start)"
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  $result = & $Block
  $sw.Stop()
  Write-Step "$Name (done) - $([math]::Round($sw.Elapsed.TotalSeconds,1))s"
  return $result
}

function Join-OrEmpty {
  param([object[]]$Values)
  if (-not $Values -or $Values.Count -eq 0) { return "" }
  ($Values | Where-Object { $_ -ne $null -and "$_" -ne "" } | ForEach-Object { "$_" }) -join ";"
}

function Resolve-Ids {
  param([string[]]$Ids,[hashtable]$Map)
  if (-not $Ids -or $Ids.Count -eq 0) { return "" }

  ($Ids | ForEach-Object {
    $id = $_
    if ([string]::IsNullOrWhiteSpace($id)) { return "" }
    if ($Map.ContainsKey($id)) { $Map[$id] } else { $id }
  } | Where-Object { $_ -ne "" }) -join ";"
}

function Resolve-UserIds {
  param([string[]]$Ids,[hashtable]$Map)
  if (-not $Ids -or $Ids.Count -eq 0) { return "" }

  ($Ids | ForEach-Object {
    $id = $_
    if ([string]::IsNullOrWhiteSpace($id)) { return "" }

    if ($Map.ContainsKey($id)) { return $Map[$id] }

    try {
      $u = Get-MgUser -UserId $id -Property DisplayName -ErrorAction Stop
      if ($u -and $u.DisplayName) {
        $Map[$id] = $u.DisplayName
        return $u.DisplayName
      }
      return $id
    } catch {
      return $id
    }
  } | Where-Object { $_ -ne "" }) -join ";"
}

function Resolve-Apps {
  param([string[]]$Ids,[hashtable]$AppMap)
  if (-not $Ids -or $Ids.Count -eq 0) { return "" }

  ($Ids | ForEach-Object {
    $id = $_
    if ([string]::IsNullOrWhiteSpace($id)) { return "" }
    if ($id -eq "All") { return "All cloud apps" }
    if ($AppMap.ContainsKey($id)) { return $AppMap[$id] }
    return $id
  } | Where-Object { $_ -ne "" }) -join ";"
}

function Safe-PropNames {
  param([object]$Obj)
  if (-not $Obj) { return "" }
  (($Obj.PSObject.Properties.Name) | Sort-Object) -join ";"
}

# --------------------------
# Main
# --------------------------
$globalSw = [System.Diagnostics.Stopwatch]::StartNew()
Write-Step "Starting export"

# 0) Verify prerequisites
if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
  throw "Microsoft.Graph module is not installed. Run prerequisites.ps1 first."
}
Import-Module Microsoft.Graph

# 1) Connect to Graph
$scopes = @("Policy.Read.All","Directory.Read.All")
Measure-Step "Connect to Microsoft Graph" {
  Connect-MgGraph -Scopes $scopes | Out-Null
  Get-MgContext | Out-Null
}

# 2) Fetch policies
$policies = Measure-Step "Fetch Conditional Access policies" {
  Get-MgIdentityConditionalAccessPolicy -All
}
Write-Step "Policies: $($policies.Count)"

# 3) Fetch directory objects for resolution
$users = Measure-Step "Fetch users (Id, DisplayName)" {
  Get-MgUser -All -Property Id,DisplayName
}
Write-Step "Users fetched: $($users.Count)"

$groups = Measure-Step "Fetch groups (Id, DisplayName)" {
  Get-MgGroup -All -Property Id,DisplayName
}
Write-Step "Groups fetched: $($groups.Count)"

$roleTemplates = Measure-Step "Fetch role templates (Id, DisplayName)" {
  Get-MgDirectoryRoleTemplate -All -Property Id,DisplayName
}
Write-Step "Role templates fetched: $($roleTemplates.Count)"

$directoryRoles = Measure-Step "Fetch active directory roles (Id, DisplayName, RoleTemplateId)" {
  Get-MgDirectoryRole -All -Property Id,DisplayName,RoleTemplateId
}
Write-Step "Active roles fetched: $($directoryRoles.Count)"

$servicePrincipals = Measure-Step "Fetch service principals (apps) (Id, AppId, DisplayName)" {
  Get-MgServicePrincipal -All -Property Id,AppId,DisplayName
}
Write-Step "Service principals fetched: $($servicePrincipals.Count)"

# 4) Build maps
$userMap  = @{}
$groupMap = @{}
$roleMap  = @{}
$appMap   = @{}

Measure-Step "Build lookup maps (users, groups, roles, apps)" {
  $users  | ForEach-Object { if ($_.Id) { $userMap[$_.Id]  = $_.DisplayName } }
  $groups | ForEach-Object { if ($_.Id) { $groupMap[$_.Id] = $_.DisplayName } }

  $roleTemplates | ForEach-Object { if ($_.Id) { $roleMap[$_.Id] = $_.DisplayName } }
  $directoryRoles | ForEach-Object {
    if ($_.Id) { $roleMap[$_.Id] = $_.DisplayName }
    if ($_.RoleTemplateId -and -not $roleMap.ContainsKey($_.RoleTemplateId)) {
      $roleMap[$_.RoleTemplateId] = $_.DisplayName
    }
  }

  $servicePrincipals | ForEach-Object {
    if ($_.Id)   { $appMap[$_.Id] = $_.DisplayName }
    if ($_.AppId){ $appMap[$_.AppId] = $_.DisplayName }
  }
}

# 5) Build overview
$overview = Measure-Step "Build overview rows" {
  $policies | Select-Object `
    Id,
    DisplayName,
    State,
    @{n="IncludeUsers";e={ Resolve-UserIds $_.Conditions.Users.IncludeUsers $userMap }},
    @{n="IncludeGroups";e={ Resolve-Ids $_.Conditions.Users.IncludeGroups $groupMap }},
    @{n="IncludeRoles";e={ Resolve-Ids $_.Conditions.Users.IncludeRoles $roleMap }},
    @{n="ExcludeUsers";e={ Resolve-UserIds $_.Conditions.Users.ExcludeUsers $userMap }},
    @{n="ExcludeGroups";e={ Resolve-Ids $_.Conditions.Users.ExcludeGroups $groupMap }},
    @{n="ExcludeRoles";e={ Resolve-Ids $_.Conditions.Users.ExcludeRoles $roleMap }},
    @{n="IncludeApps";e={ Resolve-Apps $_.Conditions.Applications.IncludeApplications $appMap }},
    @{n="ExcludeApps";e={ Resolve-Apps $_.Conditions.Applications.ExcludeApplications $appMap }},
    @{n="UserActions";e={ Join-OrEmpty $_.Conditions.Applications.IncludeUserActions }},
    @{n="ClientAppTypes";e={ Join-OrEmpty $_.Conditions.ClientAppTypes }},
    @{n="Platforms";e={ Join-OrEmpty $_.Conditions.Platforms.IncludePlatforms }},
    @{n="Locations";e={ Join-OrEmpty $_.Conditions.Locations.IncludeLocations }},
    @{n="ExcludeLocations";e={ Join-OrEmpty $_.Conditions.Locations.ExcludeLocations }},
    @{n="SignInRiskLevels";e={ Join-OrEmpty $_.Conditions.SignInRiskLevels }},
    @{n="UserRiskLevels";e={ Join-OrEmpty $_.Conditions.UserRiskLevels }},
    @{n="GrantControls";e={ Join-OrEmpty $_.GrantControls.BuiltInControls }},
    @{n="AuthStrength";e={ $_.GrantControls.AuthenticationStrength.DisplayName }},
    @{n="TermsOfUse";e={ Join-OrEmpty $_.GrantControls.TermsOfUse }},
    @{n="SessionControls";e={ Safe-PropNames $_.SessionControls }}
}

# 6) Export CSV
$path = Join-Path (Get-Location) "ConditionalAccess-overview-readable.csv"
Measure-Step "Export CSV" {
  $overview | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
}

Write-Step "Exported: $path"

# 7) Open file (optional)
Measure-Step "Open CSV (optional)" {
  try {
    if ($IsMacOS) { open $path }
    elseif (-not $IsLinux) { Invoke-Item $path }
    else { Write-Step "Linux: open the CSV manually: $path" }
  } catch {
    Write-Step "Could not auto-open the file. Path: $path"
  }
}

$globalSw.Stop()
Write-Step "Finished - total $([math]::Round($globalSw.Elapsed.TotalSeconds,1))s"