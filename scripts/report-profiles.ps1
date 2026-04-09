Set-StrictMode -Version Latest

function Resolve-ReportProfilePath {
  param(
    [AllowNull()]
    [string]$ProfileName = "experiment-report",

    [AllowNull()]
    [string]$ProfilePath,

    [AllowNull()]
    [string]$RepoRoot
  )

  if (-not [string]::IsNullOrWhiteSpace($ProfilePath)) {
    return (Resolve-Path -LiteralPath $ProfilePath).Path
  }

  if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
  }

  if ([string]::IsNullOrWhiteSpace($ProfileName)) {
    $ProfileName = "experiment-report"
  }

  return [System.IO.Path]::GetFullPath((Join-Path $RepoRoot ("profiles\{0}.json" -f $ProfileName)))
}

function Get-ReportProfile {
  param(
    [AllowNull()]
    [string]$ProfileName = "experiment-report",

    [AllowNull()]
    [string]$ProfilePath,

    [AllowNull()]
    [string]$RepoRoot
  )

  $resolvedProfilePath = Resolve-ReportProfilePath -ProfileName $ProfileName -ProfilePath $ProfilePath -RepoRoot $RepoRoot
  if (-not (Test-Path -LiteralPath $resolvedProfilePath -PathType Leaf)) {
    throw "Report profile was not found: $resolvedProfilePath"
  }

  $raw = Get-Content -LiteralPath $resolvedProfilePath -Raw -Encoding UTF8
  if ([string]::IsNullOrWhiteSpace($raw)) {
    throw "Report profile is empty: $resolvedProfilePath"
  }

  $profile = $raw | ConvertFrom-Json
  if ($null -eq $profile) {
    throw "Report profile did not parse: $resolvedProfilePath"
  }
  if (-not ($profile.PSObject.Properties.Name -contains "metadataFields") -or @($profile.metadataFields).Count -eq 0) {
    throw "Report profile is missing metadataFields: $resolvedProfilePath"
  }
  if (-not ($profile.PSObject.Properties.Name -contains "sectionFields") -or @($profile.sectionFields).Count -eq 0) {
    throw "Report profile is missing sectionFields: $resolvedProfilePath"
  }
  if (-not ($profile.PSObject.Properties.Name -contains "detailProfiles")) {
    throw "Report profile is missing detailProfiles: $resolvedProfilePath"
  }

  Add-Member -InputObject $profile -MemberType NoteProperty -Name resolvedProfilePath -Value $resolvedProfilePath -Force
  if (-not ($profile.PSObject.Properties.Name -contains "name") -or [string]::IsNullOrWhiteSpace([string]$profile.name)) {
    Add-Member -InputObject $profile -MemberType NoteProperty -Name name -Value ([System.IO.Path]::GetFileNameWithoutExtension($resolvedProfilePath)) -Force
  }

  return $profile
}

function Get-ReportProfileLabels {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$Profile
  )

  $labels = [ordered]@{}

  foreach ($field in @($Profile.metadataFields)) {
    $key = [string]$field.key
    if (-not [string]::IsNullOrWhiteSpace($key)) {
      $labels[$key] = [string]$field.label
    }
  }

  foreach ($field in @($Profile.sectionFields)) {
    $key = [string]$field.key
    if (-not [string]::IsNullOrWhiteSpace($key)) {
      $labels[$key] = [string]$field.heading
    }
  }

  foreach ($field in @($Profile.extraLabels)) {
    $key = [string]$field.key
    if (-not [string]::IsNullOrWhiteSpace($key)) {
      $labels[$key] = [string]$field.label
    }
  }

  return $labels
}

function Get-ReportProfileSectionFields {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$Profile
  )

  return @($Profile.sectionFields)
}

function Get-ReportProfileRequiredHeadings {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$Profile
  )

  return @(
    Get-ReportProfileSectionFields -Profile $Profile |
      ForEach-Object { [string]$_.heading } |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  )
}

function Get-ReportProfileDetailProfile {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$Profile,

    [Parameter(Mandatory = $true)]
    [string]$DetailLevel
  )

  if (-not ($Profile.detailProfiles.PSObject.Properties.Name -contains $DetailLevel)) {
    throw "Report profile '$([string]$Profile.name)' does not define detailProfiles.$DetailLevel"
  }

  return $Profile.detailProfiles.$DetailLevel
}
