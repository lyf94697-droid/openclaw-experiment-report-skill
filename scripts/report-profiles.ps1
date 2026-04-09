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

  foreach ($field in @(Get-ReportProfileOptionalPropertyValue -Object $Profile -Name "extraLabels")) {
    $key = [string]$field.key
    if (-not [string]::IsNullOrWhiteSpace($key)) {
      $labels[$key] = [string]$field.label
    }
  }

  return $labels
}

function ConvertTo-ReportProfilePlainHashtable {
  param(
    [AllowNull()]
    [object]$InputObject
  )

  if ($null -eq $InputObject) {
    return @{}
  }

  $table = @{}
  if ($InputObject -is [System.Collections.IDictionary]) {
    foreach ($key in $InputObject.Keys) {
      $table[[string]$key] = $InputObject[$key]
    }
    return $table
  }

  foreach ($property in $InputObject.PSObject.Properties) {
    $table[[string]$property.Name] = $property.Value
  }

  return $table
}

function ConvertTo-ReportProfileStringArray {
  param(
    [AllowNull()]
    [object]$Value
  )

  if ($null -eq $Value) {
    return @()
  }

  if (($Value -is [System.Collections.IEnumerable]) -and ($Value -isnot [string])) {
    return @(
      @($Value) |
        ForEach-Object { [string]$_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
  }

  if ([string]::IsNullOrWhiteSpace([string]$Value)) {
    return @()
  }

  return @([string]$Value)
}

function Get-ReportProfileOptionalPropertyValue {
  param(
    [AllowNull()]
    [object]$Object,

    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  if ($null -eq $Object) {
    return $null
  }

  $property = $Object.PSObject.Properties[$Name]
  if ($null -eq $property) {
    return $null
  }

  return $property.Value
}

function Get-ReportProfileSectionIdFromKey {
  param(
    [AllowNull()]
    [string]$Key
  )

  switch ([string]$Key) {
    "Purpose" { return "purpose" }
    "Environment" { return "environment" }
    "Theory" { return "theory" }
    "Steps" { return "steps" }
    "Results" { return "result" }
    "Analysis" { return "analysis" }
    "Summary" { return "summary" }
    default {
      if ([string]::IsNullOrWhiteSpace($Key)) {
        return ""
      }
      return $Key.Trim().ToLowerInvariant()
    }
  }
}

function Get-ReportProfileSectionFields {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$Profile
  )

  return @($Profile.sectionFields)
}

function Get-ReportProfileSectionRules {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$Profile
  )

  return @(
    @(Get-ReportProfileSectionFields -Profile $Profile) |
      ForEach-Object {
        $sectionId = Get-ReportProfileSectionIdFromKey -Key ([string]$_.key)
        $heading = [string]$_.heading
        $headingAliases = @(
          @($_.aliases) |
            ForEach-Object { [string]$_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )
        if ($headingAliases.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($heading)) {
          $headingAliases = @($heading)
        }

        $inputAliases = @(
          @($headingAliases + @($heading, [string]$_.key, $sectionId)) |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            Select-Object -Unique
        )

        [pscustomobject]@{
          id = $sectionId
          key = [string]$_.key
          canonicalLabel = $heading
          headingAliases = $headingAliases
          inputAliases = $inputAliases
        }
      }
  )
}

function Get-ReportProfileMetadataPrefixes {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$Profile
  )

  return @(
    @($Profile.metadataFields + @(Get-ReportProfileOptionalPropertyValue -Object $Profile -Name "extraLabels")) |
      ForEach-Object { [string]$_.label } |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
      Select-Object -Unique
  )
}

function Get-ReportProfileExtraSectionHeadings {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$Profile
  )

  return @(
    ConvertTo-ReportProfileStringArray -Value (Get-ReportProfileOptionalPropertyValue -Object $Profile -Name "extraSectionHeadings") |
      Select-Object -Unique
  )
}

function Get-ReportProfileDefaultStyleProfile {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$Profile,

    [string]$Fallback = "auto"
  )

  $resolvedFallback = if ([string]::IsNullOrWhiteSpace($Fallback)) { "auto" } else { $Fallback.Trim().ToLowerInvariant() }
  $rawValue = [string](Get-ReportProfileOptionalPropertyValue -Object $Profile -Name "defaultStyleProfile")
  if ([string]::IsNullOrWhiteSpace($rawValue)) {
    return $resolvedFallback
  }

  $normalized = $rawValue.Trim().ToLowerInvariant()
  if (@("auto", "default", "compact", "school") -notcontains $normalized) {
    throw "Report profile '$([string]$Profile.name)' has unsupported defaultStyleProfile '$rawValue'"
  }

  return $normalized
}

function Get-ReportProfileImagePlacementDefaults {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$Profile
  )

  $imagePlacementDefaults = ConvertTo-ReportProfilePlainHashtable -InputObject (Get-ReportProfileOptionalPropertyValue -Object $Profile -Name "imagePlacementDefaults")
  $defaultCaptions = ConvertTo-ReportProfilePlainHashtable -InputObject $imagePlacementDefaults["defaultCaptions"]

  return [pscustomobject]@{
    fallbackSectionOrder = @(
      ConvertTo-ReportProfileStringArray -Value $imagePlacementDefaults["fallbackSectionOrder"] |
        ForEach-Object { Get-ReportProfileSectionIdFromKey -Key ([string]$_) } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
    )
    defaultCaptions = $defaultCaptions
    filenameCaptionRules = @($imagePlacementDefaults["filenameCaptionRules"])
  }
}

function Normalize-ReportProfileLookupKey {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return ""
  }

  $normalized = $Text.ToLowerInvariant()
  $normalized = $normalized -replace '\s+', ''
  $normalized = $normalized -replace '[\p{P}\p{S}\uFF3F]+', ''
  return $normalized
}

function Get-ReportProfileImageFallbackSectionOrder {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$Profile
  )

  $configuredOrder = @(Get-ReportProfileImagePlacementDefaults -Profile $Profile).fallbackSectionOrder
  if ($configuredOrder.Count -gt 0) {
    return $configuredOrder
  }

  return @("steps", "result", "environment", "analysis", "summary", "purpose")
}

function Get-ReportProfileDefaultImageCaptionBody {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$Profile,

    [AllowNull()]
    [string]$SectionId,

    [AllowNull()]
    [string]$BaseName
  )

  $imagePlacementDefaults = Get-ReportProfileImagePlacementDefaults -Profile $Profile
  $normalizedBaseName = Normalize-ReportProfileLookupKey -Text $BaseName

  foreach ($rule in @($imagePlacementDefaults.filenameCaptionRules)) {
    $ruleTable = ConvertTo-ReportProfilePlainHashtable -InputObject $rule
    $pattern = [string]$ruleTable["pattern"]
    $caption = [string]$ruleTable["caption"]
    if (-not [string]::IsNullOrWhiteSpace($pattern) -and -not [string]::IsNullOrWhiteSpace($caption) -and $normalizedBaseName -match $pattern) {
      return $caption
    }
  }

  $defaultCaptions = $imagePlacementDefaults.defaultCaptions
  if (-not [string]::IsNullOrWhiteSpace($SectionId) -and $defaultCaptions.ContainsKey($SectionId) -and -not [string]::IsNullOrWhiteSpace([string]$defaultCaptions[$SectionId])) {
    return [string]$defaultCaptions[$SectionId]
  }
  if ($defaultCaptions.ContainsKey("default") -and -not [string]::IsNullOrWhiteSpace([string]$defaultCaptions["default"])) {
    return [string]$defaultCaptions["default"]
  }

  switch ([string]$SectionId) {
    "environment" { return "实验环境截图" }
    "steps" { return "实验步骤截图" }
    "result" { return "实验结果截图" }
    "analysis" { return "问题分析截图" }
    "summary" { return "实验总结截图" }
    default { return "实验过程截图" }
  }
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
