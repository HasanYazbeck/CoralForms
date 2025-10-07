<# 
.SYNOPSIS
  Export specified SharePoint Online lists as templates with content and download them locally.

.DESCRIPTION
  - Ensures Custom Script is enabled on the site collection and the current web (subsite).
  - Tries to save each list as a .stp template (with content) to the List Template Gallery.
  - Downloads the .stp to a local folder.
  - If .stp save isn’t allowed/supported, falls back to a PnP provisioning package (.pnp) including content.

.REQUIREMENTS
  - PowerShell 5.1+
  - PnP.PowerShell module (Install-Module PnP.PowerShell -Scope CurrentUser)
  - Permissions: Site Collection Admin (recommended) or at least site owner + ability to enable custom script.
.One-time set up:
  -Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
  -Install-Module PnP.PowerShell -Scope CurrentUser

.PARAMETER SiteUrl
  URL of the SharePoint Online subsite (or site) containing the lists.

.PARAMETER ListTitles
  Array of list titles to export.

.PARAMETER OutputFolder
  Local folder path to store exported files.

.PARAMETER Interactive
  Use interactive login (recommended).

.EXAMPLE
  .\Export-SPOListsAsTemplates.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Team/SubA" `
    -ListTitles @("PPEForm","PPEFormItems","PPE_Form_Approval_Workflow") `
    -OutputFolder "C:\SPFx_Online_Solutions\Coral_Lebanon\CoralForms\PowerShell Scripts\ExportSharePointLists" `
    -Interactive
#>

[CmdletBinding()]
param (
  [Parameter(Mandatory = $true)]
  [string]$SiteUrl,

  [Parameter(Mandatory = $true)]
  [string]$OutputFolder,

  [switch]$Interactive
)

$ListTitles = @(
  "PPEForm",
  "PPEFormItems",
  "PPE_Form_Approval_Workflow"
)

if (-not $ListTitles -or $ListTitles.Count -eq 0) {
  Write-Error "No list titles configured. Please edit the script and add list names to `$ListTitles."
  exit 1
}

function Ensure-Module {
  param([string]$Name)
  if (-not (Get-Module -ListAvailable -Name $Name)) {
    Write-Host "Installing module $Name ..." -ForegroundColor Yellow
    Install-Module $Name -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module $Name -ErrorAction Stop
}

function New-SafeFileName {
  param([string]$Base, [string]$Ext)
  $safeBase = ($Base -replace '[^\w\-]+','_').Trim('_')
  return "{0}-{1}{2}" -f $safeBase, (Get-Date -Format 'yyyyMMdd-HHmmss'), $Ext
}

function Ensure-CustomScript {
  param(
    [Microsoft.SharePoint.Client.ClientContext]$CtxWeb,
    [object]$ConnWeb,
    [object]$ConnRoot
  )
  try {
    # Site collection level (DenyAddAndCustomizePages = false)
    $rootUrl = $CtxWeb.Site.Url
    Write-Host "Ensuring Custom Script is enabled on site collection: $rootUrl" -ForegroundColor Cyan
    Set-PnPSite -Identity $rootUrl -DenyAddAndCustomizePages:$false -Connection $ConnRoot -ErrorAction SilentlyContinue

    # Web level (NoScriptSite = false)
    Write-Host "Ensuring Custom Script is enabled on web: $($CtxWeb.Url)" -ForegroundColor Cyan
    Set-PnPWeb -NoScriptSite:$false -Connection $ConnWeb -ErrorAction SilentlyContinue

    # Poll a few times; in SPO this flag might take time to flip
    $max = 10; $sleep = 3; $ok = $false
    for ($i=1; $i -le $max; $i++) {
      Start-Sleep -Seconds $sleep
      $w = Get-PnPWeb -Includes NoScriptSite -Connection $ConnWeb
      if (-not $w.NoScriptSite) { $ok = $true; break }
    }
    if (-not $ok) {
      Write-Warning "Custom Script may still be propagating. SaveAsTemplate could fail; fallback to PnP template will be attempted if needed."
    }
  } catch {
    Write-Warning "Failed to verify/enable Custom Script: $($_.Exception.Message)"
  }
}

function Save-ListAsStpAndDownload {
  param(
    [string]$ListTitle,
    [Microsoft.SharePoint.Client.ClientContext]$CtxWeb,
    [object]$ConnWeb,
    [object]$ConnRoot,
    [string]$OutputFolder
  )

  # Ensure list exists
  $list = $null
  try {
    $list = Get-PnPList -Identity $ListTitle -Connection $ConnWeb -ErrorAction Stop
  } catch {
    Write-Warning "List '$ListTitle' not found on $($CtxWeb.Url). Skipping."
    return $false
  }

  $fileName = New-SafeFileName -Base $ListTitle -Ext ".stp"
  $templateTitle = $ListTitle
  $templateDesc = "Exported by script on $(Get-Date -Format 'u')"

  try {
    Write-Host "Saving '$ListTitle' as template (with content) -> $fileName" -ForegroundColor Green
    # CSOM SaveAsTemplate (uploads to site collection List Template Gallery)
    $web = $CtxWeb.Web
    $spList = $web.Lists.GetByTitle($ListTitle)
    $CtxWeb.Load($spList)
    $CtxWeb.ExecuteQuery()
    $spList.SaveAsTemplate($fileName, $templateTitle, $templateDesc, $true)
    $CtxWeb.ExecuteQuery()

    # File should be at the site collection's List Template Gallery
    $rootRel = $CtxWeb.Site.RootWeb.ServerRelativeUrl
    if ([string]::IsNullOrEmpty($rootRel) -or $rootRel -eq "/") {
      $ltRelUrl = "/_catalogs/lt/$fileName"
    } else {
      $ltRelUrl = "$rootRel/_catalogs/lt/$fileName"
    }

    # Wait until the file appears (poll)
    $max = 20; $sleep = 3; $found = $false
    for ($i=1; $i -le $max; $i++) {
      try {
        Get-PnPFile -Url $ltRelUrl -Connection $ConnRoot -AsFile -Path $OutputFolder -FileName $fileName -Force -ErrorAction Stop | Out-Null
        $found = $true
        break
      } catch {
        Start-Sleep -Seconds $sleep
      }
    }
    if ($found) {
      Write-Host "Downloaded: $OutputFolder\$fileName" -ForegroundColor Green
      return $true
    } else {
      throw "Template file not found at $ltRelUrl after waiting."
    }
  } catch {
    Write-Warning "SaveAsTemplate failed for '$ListTitle': $($_.Exception.Message)"
    return $false
  }
}

function Export-ListAsPnP {
  param(
    [string]$ListTitle,
    [object]$ConnWeb,
    [string]$OutputFolder
  )
  try {
    $pnpName = New-SafeFileName -Base $ListTitle -Ext ".pnp"
    $outPath = Join-Path $OutputFolder $pnpName
    Write-Host "Falling back: Exporting '$ListTitle' as PnP provisioning package with content -> $pnpName" -ForegroundColor Yellow
    Get-PnPProvisioningTemplate `
      -Connection $ConnWeb `
      -Out $outPath `
      -Handlers Lists `
      -ListsToExtract $ListTitle `
      -IncludeContent `
      -ExcludeHandlers Workflows,ApplicationLifecycleManagement,Publishing `
      -ErrorAction Stop
    Write-Host "Exported PnP template: $outPath" -ForegroundColor Green
    return $true
  } catch {
    Write-Warning "PnP export failed for '$ListTitle': $($_.Exception.Message)"
    return $false
  }
}

# -------- Main --------

Ensure-Module -Name "PnP.PowerShell"

# Make sure output folder exists
if (-not (Test-Path -Path $OutputFolder)) {
  New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

# Connect to subsite (web) and derive site collection (root) url
if ($Interactive) {
  $connWeb  = Connect-PnPOnline -Url $SiteUrl -Interactive -ReturnConnection
} else {
  # If you prefer Device Login uncomment the next line and comment Interactive above:
  # $connWeb  = Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ReturnConnection
  $connWeb  = Connect-PnPOnline -Url $SiteUrl -Interactive -ReturnConnection
}
$ctxWeb  = Get-PnPContext -Connection $connWeb
$rootUrl = $ctxWeb.Site.Url

# Reuse the same token for the site collection connection
$connRoot = Connect-PnPOnline -Url $rootUrl -ReturnConnection

# Ensure Custom Script
Ensure-CustomScript -CtxWeb $ctxWeb -ConnWeb $connWeb -ConnRoot $connRoot

# Iterate lists
$results = @()
foreach ($title in $ListTitles) {
  Write-Host "Processing list: $title" -ForegroundColor Cyan
  $ok = Save-ListAsStpAndDownload -ListTitle $title -CtxWeb $ctxWeb -ConnWeb $connWeb -ConnRoot $connRoot -OutputFolder $OutputFolder
  if (-not $ok) {
    # Fallback to PnP provisioning export with content
    $ok = Export-ListAsPnP -ListTitle $title -ConnWeb $connWeb -OutputFolder $OutputFolder
  }
  $results += [pscustomobject]@{
    ListTitle = $title
    Success   = $ok
  }
}

Write-Host "`nSummary:" -ForegroundColor Cyan
$results | Format-Table -AutoSize