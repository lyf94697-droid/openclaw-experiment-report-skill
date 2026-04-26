[CmdletBinding()]
param(
  [string]$ProfileDir,

  [string[]]$ProfilePath,

  [ValidateSet("text", "json")]
  [string]$Format = "text",

  [string]$OutFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-OptionalPropertyValue {
  param(
    [AllowNull()]
    [object]$InputObject,

    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  if ($null -eq $InputObject) {
    return $null
  }

  $property = $InputObject.PSObject.Properties[$Name]
  if ($null -eq $property) {
    return $null
  }

  return $property.Value
}

function Get-StringItems {
  param(
    [AllowNull()]
    [object]$Value
  )

  $items = New-Object System.Collections.Generic.List[string]
  if ($null -eq $Value) {
    return @()
  }

  if (($Value -is [System.Collections.IEnumerable]) -and ($Value -isnot [string])) {
    foreach ($item in $Value) {
      foreach ($text in (Get-StringItems -Value $item)) {
        if (-not [string]::IsNullOrWhiteSpace($text)) {
          [void]$items.Add($text)
        }
      }
    }
    return $items.ToArray()
  }

  $stringValue = [string]$Value
  if (-not [string]::IsNullOrWhiteSpace($stringValue)) {
    [void]$items.Add($stringValue.Trim())
  }

  return $items.ToArray()
}

function Test-IntegerValue {
  param(
    [AllowNull()]
    [object]$Value
  )

  if ($null -eq $Value) {
    return $false
  }

  if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int] -or $Value -is [long]) {
    return $true
  }

  return $false
}

function Get-SectionIdFromKey {
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

function New-ProfileFinding {
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("error", "warning")]
    [string]$Severity,

    [Parameter(Mandatory = $true)]
    [string]$Code,

    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$Message
  )

  [pscustomobject]@{
    severity = $Severity
    code = $Code
    path = $Path
    message = $Message
  }
}

function Add-ProfileFinding {
  param(
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$Findings,

    [Parameter(Mandatory = $true)]
    [ValidateSet("error", "warning")]
    [string]$Severity,

    [Parameter(Mandatory = $true)]
    [string]$Code,

    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$Message
  )

  [void]$Findings.Add((New-ProfileFinding -Severity $Severity -Code $Code -Path $Path -Message $Message))
}

function Get-ResolvedProfilePaths {
  param(
    [AllowNull()]
    [string]$ProfileDirectory,

    [AllowNull()]
    [string[]]$RequestedProfilePaths,

    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$Findings
  )

  $paths = New-Object System.Collections.Generic.List[string]
  if ($null -ne $RequestedProfilePaths -and $RequestedProfilePaths.Count -gt 0) {
    foreach ($requestedPath in $RequestedProfilePaths) {
      if ([string]::IsNullOrWhiteSpace($requestedPath)) {
        continue
      }
      if (-not (Test-Path -LiteralPath $requestedPath)) {
        Add-ProfileFinding -Findings $Findings -Severity error -Code "profile-file-missing" -Path $requestedPath -Message "Profile path does not exist."
        continue
      }
      $resolvedPath = (Resolve-Path -LiteralPath $requestedPath).Path
      [void]$paths.Add($resolvedPath)
    }
    return $paths.ToArray()
  }

  if ([string]::IsNullOrWhiteSpace($ProfileDirectory)) {
    $repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
    $ProfileDirectory = Join-Path $repoRoot "profiles"
  }

  if (-not (Test-Path -LiteralPath $ProfileDirectory)) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "profile-dir-missing" -Path $ProfileDirectory -Message "Profile directory does not exist."
    return @()
  }

  return @(
    Get-ChildItem -LiteralPath $ProfileDirectory -Filter "*.json" -File |
      Where-Object { $_.Name -ne "report-profile.schema.json" } |
      Sort-Object -Property Name |
      ForEach-Object { $_.FullName }
  )
}

function Read-ProfileJson {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$Findings
  )

  try {
    $jsonText = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($jsonText)) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "profile-json-empty" -Path $Path -Message "Profile JSON file is empty."
      return $null
    }
    return ($jsonText | ConvertFrom-Json)
  } catch {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "profile-json-parse-error" -Path $Path -Message $_.Exception.Message
    return $null
  }
}

function Assert-NonEmptyString {
  param(
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$Findings,

    [AllowNull()]
    [object]$Value,

    [Parameter(Mandatory = $true)]
    [string]$Code,

    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$Message
  )

  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code $Code -Path $Path -Message $Message
    return $false
  }

  return $true
}

function Test-FieldList {
  param(
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$Findings,

    [AllowNull()]
    [object]$Fields,

    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$CodePrefix,

    [switch]$Required
  )

  $items = @($Fields)
  if ($Required -and ($null -eq $Fields -or $items.Count -eq 0)) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "$CodePrefix-empty" -Path $Path -Message "Field list must contain at least one item."
    return @()
  }

  $seenKeys = @{}
  for ($index = 0; $index -lt $items.Count; $index++) {
    $itemPath = "$Path[$index]"
    $key = [string](Get-OptionalPropertyValue -InputObject $items[$index] -Name "key")
    $label = [string](Get-OptionalPropertyValue -InputObject $items[$index] -Name "label")
    [void](Assert-NonEmptyString -Findings $Findings -Value $key -Code "$CodePrefix-key-empty" -Path "$itemPath.key" -Message "Field key must not be empty.")
    [void](Assert-NonEmptyString -Findings $Findings -Value $label -Code "$CodePrefix-label-empty" -Path "$itemPath.label" -Message "Field label must not be empty.")

    if (-not [string]::IsNullOrWhiteSpace($key)) {
      if ($key -notmatch "^[A-Za-z][A-Za-z0-9_]*$") {
        Add-ProfileFinding -Findings $Findings -Severity error -Code "$CodePrefix-key-invalid" -Path "$itemPath.key" -Message "Field key must be an ASCII identifier."
      }
      $normalizedKey = $key.ToLowerInvariant()
      if ($seenKeys.ContainsKey($normalizedKey)) {
        Add-ProfileFinding -Findings $Findings -Severity error -Code "$CodePrefix-key-duplicate" -Path "$itemPath.key" -Message "Field key '$key' is duplicated."
      } else {
        $seenKeys[$normalizedKey] = $true
      }
    }

    $aliases = Get-OptionalPropertyValue -InputObject $items[$index] -Name "aliases"
    if ($null -ne $aliases) {
      $rawAliasCount = @($aliases).Count
      $cleanAliases = @(Get-StringItems -Value $aliases)
      if ($cleanAliases.Count -ne $rawAliasCount) {
        Add-ProfileFinding -Findings $Findings -Severity error -Code "$CodePrefix-alias-empty" -Path "$itemPath.aliases" -Message "Aliases must not contain empty values."
      }
    }
  }

  return $items
}

function Test-MinCharsPair {
  param(
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$Findings,

    [AllowNull()]
    [object]$MinChars,

    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  if ($null -eq $MinChars) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "section-minchars-missing" -Path $Path -Message "Section minChars must include standard and full values."
    return
  }

  $standard = Get-OptionalPropertyValue -InputObject $MinChars -Name "standard"
  $full = Get-OptionalPropertyValue -InputObject $MinChars -Name "full"

  if (-not (Test-IntegerValue -Value $standard)) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "section-minchars-standard-invalid" -Path "$Path.standard" -Message "standard minChars must be an integer."
  }
  if (-not (Test-IntegerValue -Value $full)) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "section-minchars-full-invalid" -Path "$Path.full" -Message "full minChars must be an integer."
  }
  if ((Test-IntegerValue -Value $standard) -and (Test-IntegerValue -Value $full) -and ([int]$full -lt [int]$standard)) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "section-minchars-order-invalid" -Path $Path -Message "full minChars must be greater than or equal to standard minChars."
  }
}

function Test-SectionFields {
  param(
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$Findings,

    [AllowNull()]
    [object]$SectionFields,

    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $items = @($SectionFields)
  if ($null -eq $SectionFields -or $items.Count -eq 0) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "section-fields-empty" -Path $Path -Message "sectionFields must contain at least one item."
    return @{}
  }

  $seenKeys = @{}
  $seenHeadings = @{}
  $sectionIds = @{}

  for ($index = 0; $index -lt $items.Count; $index++) {
    $itemPath = "$Path[$index]"
    $key = [string](Get-OptionalPropertyValue -InputObject $items[$index] -Name "key")
    $heading = [string](Get-OptionalPropertyValue -InputObject $items[$index] -Name "heading")
    [void](Assert-NonEmptyString -Findings $Findings -Value $key -Code "section-key-empty" -Path "$itemPath.key" -Message "Section key must not be empty.")
    [void](Assert-NonEmptyString -Findings $Findings -Value $heading -Code "section-heading-empty" -Path "$itemPath.heading" -Message "Section heading must not be empty.")

    if (-not [string]::IsNullOrWhiteSpace($key)) {
      if ($key -notmatch "^[A-Za-z][A-Za-z0-9_]*$") {
        Add-ProfileFinding -Findings $Findings -Severity error -Code "section-key-invalid" -Path "$itemPath.key" -Message "Section key must be an ASCII identifier."
      }
      $normalizedKey = $key.ToLowerInvariant()
      if ($seenKeys.ContainsKey($normalizedKey)) {
        Add-ProfileFinding -Findings $Findings -Severity error -Code "section-key-duplicate" -Path "$itemPath.key" -Message "Section key '$key' is duplicated."
      } else {
        $seenKeys[$normalizedKey] = $true
      }

      $sectionId = Get-SectionIdFromKey -Key $key
      if (-not [string]::IsNullOrWhiteSpace($sectionId)) {
        $sectionIds[$sectionId] = $true
      }
    }

    if (-not [string]::IsNullOrWhiteSpace($heading)) {
      $normalizedHeading = $heading.Trim()
      if ($seenHeadings.ContainsKey($normalizedHeading)) {
        Add-ProfileFinding -Findings $Findings -Severity error -Code "section-heading-duplicate" -Path "$itemPath.heading" -Message "Section heading '$heading' is duplicated."
      } else {
        $seenHeadings[$normalizedHeading] = $true
      }
    }

    $aliases = Get-OptionalPropertyValue -InputObject $items[$index] -Name "aliases"
    $cleanAliases = @(Get-StringItems -Value $aliases)
    if ($null -eq $aliases -or $cleanAliases.Count -eq 0) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "section-aliases-empty" -Path "$itemPath.aliases" -Message "Section aliases must contain at least one item."
    } elseif ($cleanAliases.Count -ne @($aliases).Count) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "section-alias-empty" -Path "$itemPath.aliases" -Message "Section aliases must not contain empty values."
    }

    Test-MinCharsPair -Findings $Findings -MinChars (Get-OptionalPropertyValue -InputObject $items[$index] -Name "minChars") -Path "$itemPath.minChars"
  }

  return $sectionIds
}

function Test-DetailProfiles {
  param(
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$Findings,

    [AllowNull()]
    [object]$DetailProfiles,

    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  if ($null -eq $DetailProfiles) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "detail-profiles-missing" -Path $Path -Message "detailProfiles must define standard and full."
    return
  }

  $minCharsByLevel = @{}
  foreach ($level in @("standard", "full")) {
    $detailProfile = Get-OptionalPropertyValue -InputObject $DetailProfiles -Name $level
    $levelPath = "$Path.$level"
    if ($null -eq $detailProfile) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "detail-profile-missing" -Path $levelPath -Message "Detail profile '$level' is missing."
      continue
    }

    $minChars = Get-OptionalPropertyValue -InputObject $detailProfile -Name "minChars"
    if (-not (Test-IntegerValue -Value $minChars)) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "detail-minchars-invalid" -Path "$levelPath.minChars" -Message "Detail profile minChars must be an integer."
    } else {
      $minCharsByLevel[$level] = [int]$minChars
    }

    $promptGuidance = @(Get-StringItems -Value (Get-OptionalPropertyValue -InputObject $detailProfile -Name "promptGuidance"))
    if ($promptGuidance.Count -eq 0) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "detail-prompt-guidance-empty" -Path "$levelPath.promptGuidance" -Message "Detail profile promptGuidance must contain at least one item."
    }
  }

  if ($minCharsByLevel.ContainsKey("standard") -and $minCharsByLevel.ContainsKey("full") -and $minCharsByLevel["full"] -lt $minCharsByLevel["standard"]) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "detail-minchars-order-invalid" -Path $Path -Message "full detail minChars must be greater than or equal to standard detail minChars."
  }
}

function Test-ImagePlacementDefaults {
  param(
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$Findings,

    [AllowNull()]
    [object]$ImagePlacementDefaults,

    [Parameter(Mandatory = $true)]
    [hashtable]$KnownSectionIds,

    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  if ($null -eq $ImagePlacementDefaults) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "image-placement-defaults-missing" -Path $Path -Message "imagePlacementDefaults must be defined."
    return
  }

  $fallbackSectionOrder = @(Get-StringItems -Value (Get-OptionalPropertyValue -InputObject $ImagePlacementDefaults -Name "fallbackSectionOrder"))
  if ($fallbackSectionOrder.Count -eq 0) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "image-fallback-order-empty" -Path "$Path.fallbackSectionOrder" -Message "fallbackSectionOrder must contain at least one section id."
  }
  foreach ($sectionId in $fallbackSectionOrder) {
    $normalizedSectionId = Get-SectionIdFromKey -Key $sectionId
    if (-not $KnownSectionIds.ContainsKey($normalizedSectionId)) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "image-fallback-order-unknown-section" -Path "$Path.fallbackSectionOrder" -Message "Fallback section '$sectionId' does not reference a known section."
    }
  }

  $defaultImageWidthCm = Get-OptionalPropertyValue -InputObject $ImagePlacementDefaults -Name "defaultImageWidthCm"
  if ($null -ne $defaultImageWidthCm -and -not [string]::IsNullOrWhiteSpace([string]$defaultImageWidthCm)) {
    $parsedWidth = 0.0
    if (-not [double]::TryParse([string]$defaultImageWidthCm, [ref]$parsedWidth) -or $parsedWidth -le 0) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "image-default-width-invalid" -Path "$Path.defaultImageWidthCm" -Message "defaultImageWidthCm must be a positive number."
    }
  }

  $autoRowLayout = Get-OptionalPropertyValue -InputObject $ImagePlacementDefaults -Name "autoRowLayout"
  if ($null -ne $autoRowLayout -and -not [string]::IsNullOrWhiteSpace([string]$autoRowLayout)) {
    if (($autoRowLayout -isnot [bool]) -and (@("true", "false", "auto", "enabled", "disabled", "explicit-only") -notcontains ([string]$autoRowLayout).Trim().ToLowerInvariant())) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "image-auto-row-layout-invalid" -Path "$Path.autoRowLayout" -Message "autoRowLayout must be a boolean or one of true, false, auto, enabled, disabled, explicit-only."
    }
  }

  $defaultCaptions = Get-OptionalPropertyValue -InputObject $ImagePlacementDefaults -Name "defaultCaptions"
  if ($null -eq $defaultCaptions) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "image-default-captions-missing" -Path "$Path.defaultCaptions" -Message "defaultCaptions must be defined."
  } else {
    $defaultCaption = Get-OptionalPropertyValue -InputObject $defaultCaptions -Name "default"
    [void](Assert-NonEmptyString -Findings $Findings -Value $defaultCaption -Code "image-default-caption-empty" -Path "$Path.defaultCaptions.default" -Message "defaultCaptions.default must not be empty.")
  }

  $captionRules = Get-OptionalPropertyValue -InputObject $ImagePlacementDefaults -Name "filenameCaptionRules"
  if ($null -eq $captionRules) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "image-caption-rules-missing" -Path "$Path.filenameCaptionRules" -Message "filenameCaptionRules must be defined."
    return
  }

  $rules = @($captionRules)
  for ($index = 0; $index -lt $rules.Count; $index++) {
    $rulePath = "$Path.filenameCaptionRules[$index]"
    $pattern = [string](Get-OptionalPropertyValue -InputObject $rules[$index] -Name "pattern")
    $caption = [string](Get-OptionalPropertyValue -InputObject $rules[$index] -Name "caption")
    if (Assert-NonEmptyString -Findings $Findings -Value $pattern -Code "image-caption-pattern-empty" -Path "$rulePath.pattern" -Message "Caption rule pattern must not be empty.") {
      try {
        [void](New-Object System.Text.RegularExpressions.Regex($pattern))
      } catch {
        Add-ProfileFinding -Findings $Findings -Severity error -Code "image-caption-pattern-invalid" -Path "$rulePath.pattern" -Message ("Caption rule pattern is not a valid regex: {0}" -f $_.Exception.Message)
      }
    }
    [void](Assert-NonEmptyString -Findings $Findings -Value $caption -Code "image-caption-caption-empty" -Path "$rulePath.caption" -Message "Caption rule caption must not be empty.")
  }
}

function Test-PaginationRiskThresholds {
  param(
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$Findings,

    [AllowNull()]
    [object]$Thresholds,

    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  if ($null -eq $Thresholds) {
    return
  }

  if ($Thresholds -is [string] -or $Thresholds -is [System.ValueType] -or (($Thresholds -is [System.Collections.IEnumerable]) -and ($Thresholds -isnot [System.Collections.IDictionary]))) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "pagination-risk-thresholds-invalid" -Path $Path -Message "paginationRiskThresholds must be an object when present."
    return
  }

  $knownKeys = @("longSectionChars", "denseSectionChars", "denseSectionParagraphs", "figureClusterRefs")
  foreach ($property in $Thresholds.PSObject.Properties) {
    if ($knownKeys -notcontains [string]$property.Name) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "pagination-risk-threshold-unknown" -Path "$Path.$($property.Name)" -Message "paginationRiskThresholds contains an unknown property."
    }
  }

  foreach ($key in $knownKeys) {
    $value = Get-OptionalPropertyValue -InputObject $Thresholds -Name $key
    if ($null -eq $value) {
      continue
    }

    if (-not (Test-IntegerValue -Value $value) -or [int64]$value -lt 0) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "pagination-risk-threshold-invalid" -Path "$Path.$key" -Message "paginationRiskThresholds.$key must be a non-negative integer."
    }
  }
}

function Test-PaginationRiskRemediations {
  param(
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$Findings,

    [AllowNull()]
    [object]$Remediations,

    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  if ($null -eq $Remediations) {
    return
  }

  if ($Remediations -is [string] -or $Remediations -is [System.ValueType] -or (($Remediations -is [System.Collections.IEnumerable]) -and ($Remediations -isnot [System.Collections.IDictionary]))) {
    Add-ProfileFinding -Findings $Findings -Severity error -Code "pagination-risk-remediations-invalid" -Path $Path -Message "paginationRiskRemediations must be an object when present."
    return
  }

  $knownKeys = @("pagination-risk-long-section", "pagination-risk-dense-section-block", "pagination-risk-figure-cluster")
  foreach ($property in $Remediations.PSObject.Properties) {
    if ($knownKeys -notcontains [string]$property.Name) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "pagination-risk-remediation-unknown" -Path "$Path.$($property.Name)" -Message "paginationRiskRemediations contains an unknown property."
    }
  }

  foreach ($key in $knownKeys) {
    $rawValue = Get-OptionalPropertyValue -InputObject $Remediations -Name $key
    if ($null -eq $rawValue) {
      continue
    }

    if ([string]::IsNullOrWhiteSpace([string]$rawValue)) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "pagination-risk-remediation-empty" -Path "$Path.$key" -Message "paginationRiskRemediations.$key must not be empty."
    }
  }
}

function Test-FieldMapCompositeRules {
  param(
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$Findings,

    [AllowNull()]
    [object]$Rules,

    [Parameter(Mandatory = $true)]
    [hashtable]$KnownSectionIds,

    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  if ($null -eq $Rules) {
    return
  }

  $ruleItems = @($Rules)
  for ($ruleIndex = 0; $ruleIndex -lt $ruleItems.Count; $ruleIndex++) {
    $rulePath = "$Path[$ruleIndex]"
    $matchAll = @(Get-StringItems -Value (Get-OptionalPropertyValue -InputObject $ruleItems[$ruleIndex] -Name "matchAll"))
    if ($matchAll.Count -eq 0) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "field-map-rule-match-all-empty" -Path "$rulePath.matchAll" -Message "fieldMapCompositeRules.matchAll must contain at least one item."
    }

    $blocks = Get-OptionalPropertyValue -InputObject $ruleItems[$ruleIndex] -Name "blocks"
    if ($null -eq $blocks -or @($blocks).Count -eq 0) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "field-map-rule-blocks-empty" -Path "$rulePath.blocks" -Message "fieldMapCompositeRules.blocks must contain at least one block."
      continue
    }

    $blockItems = @($blocks)
    for ($blockIndex = 0; $blockIndex -lt $blockItems.Count; $blockIndex++) {
      $blockPath = "$rulePath.blocks[$blockIndex]"
      $heading = [string](Get-OptionalPropertyValue -InputObject $blockItems[$blockIndex] -Name "heading")
      [void](Assert-NonEmptyString -Findings $Findings -Value $heading -Code "field-map-block-heading-empty" -Path "$blockPath.heading" -Message "Composite block heading must not be empty.")

      $sectionIds = @(Get-StringItems -Value (Get-OptionalPropertyValue -InputObject $blockItems[$blockIndex] -Name "sectionIds"))
      if ($sectionIds.Count -eq 0) {
        Add-ProfileFinding -Findings $Findings -Severity error -Code "field-map-block-section-ids-empty" -Path "$blockPath.sectionIds" -Message "Composite block sectionIds must contain at least one section id."
        continue
      }

      foreach ($sectionId in $sectionIds) {
        $normalizedSectionId = Get-SectionIdFromKey -Key $sectionId
        if (-not $KnownSectionIds.ContainsKey($normalizedSectionId)) {
          Add-ProfileFinding -Findings $Findings -Severity error -Code "field-map-block-unknown-section" -Path "$blockPath.sectionIds" -Message "Composite block section id '$sectionId' does not reference a known section."
        }
      }
    }
  }
}

function Test-ReportProfile {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [object]$Profile,

    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$Findings
  )

  $profileName = [string](Get-OptionalPropertyValue -InputObject $Profile -Name "name")
  $fileBaseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)

  [void](Assert-NonEmptyString -Findings $Findings -Value $profileName -Code "profile-name-empty" -Path "$Path.name" -Message "Profile name must not be empty.")
  if (-not [string]::IsNullOrWhiteSpace($profileName)) {
    if ($profileName -notmatch "^[a-z][a-z0-9]*(?:-[a-z0-9]+)*$") {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "profile-name-invalid" -Path "$Path.name" -Message "Profile name must use lower hyphen-case."
    }
    if ($profileName -ne $fileBaseName) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "profile-name-file-mismatch" -Path "$Path.name" -Message "Profile name must match its filename."
    }
  }

  [void](Assert-NonEmptyString -Findings $Findings -Value (Get-OptionalPropertyValue -InputObject $Profile -Name "displayName") -Code "profile-display-name-empty" -Path "$Path.displayName" -Message "displayName must not be empty.")

  $styleProfile = [string](Get-OptionalPropertyValue -InputObject $Profile -Name "defaultStyleProfile")
  if (Assert-NonEmptyString -Findings $Findings -Value $styleProfile -Code "profile-default-style-empty" -Path "$Path.defaultStyleProfile" -Message "defaultStyleProfile must not be empty.") {
    if (@("auto", "default", "compact", "school", "excellent") -notcontains $styleProfile.Trim().ToLowerInvariant()) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "profile-default-style-unsupported" -Path "$Path.defaultStyleProfile" -Message "defaultStyleProfile must be auto, default, compact, school, or excellent."
    }
  }

  $metadataFields = Test-FieldList -Findings $Findings -Fields (Get-OptionalPropertyValue -InputObject $Profile -Name "metadataFields") -Path "$Path.metadataFields" -CodePrefix "metadata-field" -Required
  $metadataKeys = @($metadataFields | ForEach-Object { [string](Get-OptionalPropertyValue -InputObject $_ -Name "key") })
  foreach ($requiredKey in @("CourseName", "ExperimentName")) {
    if ($metadataKeys -notcontains $requiredKey) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "metadata-required-key-missing" -Path "$Path.metadataFields" -Message "metadataFields must include '$requiredKey'."
    }
  }

  [void](Test-FieldList -Findings $Findings -Fields (Get-OptionalPropertyValue -InputObject $Profile -Name "extraLabels") -Path "$Path.extraLabels" -CodePrefix "extra-label")

  $knownSectionIds = Test-SectionFields -Findings $Findings -SectionFields (Get-OptionalPropertyValue -InputObject $Profile -Name "sectionFields") -Path "$Path.sectionFields"

  $extraSectionHeadings = Get-OptionalPropertyValue -InputObject $Profile -Name "extraSectionHeadings"
  if ($null -ne $extraSectionHeadings) {
    $rawCount = @($extraSectionHeadings).Count
    $cleanCount = @(Get-StringItems -Value $extraSectionHeadings).Count
    if ($cleanCount -ne $rawCount) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "extra-section-heading-empty" -Path "$Path.extraSectionHeadings" -Message "extraSectionHeadings must not contain empty values."
    }
  }

  Test-ImagePlacementDefaults -Findings $Findings -ImagePlacementDefaults (Get-OptionalPropertyValue -InputObject $Profile -Name "imagePlacementDefaults") -KnownSectionIds $knownSectionIds -Path "$Path.imagePlacementDefaults"
  Test-PaginationRiskThresholds -Findings $Findings -Thresholds (Get-OptionalPropertyValue -InputObject $Profile -Name "paginationRiskThresholds") -Path "$Path.paginationRiskThresholds"
  Test-PaginationRiskRemediations -Findings $Findings -Remediations (Get-OptionalPropertyValue -InputObject $Profile -Name "paginationRiskRemediations") -Path "$Path.paginationRiskRemediations"
  Test-FieldMapCompositeRules -Findings $Findings -Rules (Get-OptionalPropertyValue -InputObject $Profile -Name "fieldMapCompositeRules") -KnownSectionIds $knownSectionIds -Path "$Path.fieldMapCompositeRules"
  Test-FieldMapCompositeRules -Findings $Findings -Rules (Get-OptionalPropertyValue -InputObject $Profile -Name "paragraphCompositeRules") -KnownSectionIds $knownSectionIds -Path "$Path.paragraphCompositeRules"
  Test-DetailProfiles -Findings $Findings -DetailProfiles (Get-OptionalPropertyValue -InputObject $Profile -Name "detailProfiles") -Path "$Path.detailProfiles"

  $forbiddenPatterns = Get-OptionalPropertyValue -InputObject $Profile -Name "forbiddenPatterns"
  if ($null -ne $forbiddenPatterns) {
    $rawCount = @($forbiddenPatterns).Count
    $cleanCount = @(Get-StringItems -Value $forbiddenPatterns).Count
    if ($cleanCount -ne $rawCount) {
      Add-ProfileFinding -Findings $Findings -Severity error -Code "forbidden-pattern-empty" -Path "$Path.forbiddenPatterns" -Message "forbiddenPatterns must not contain empty values."
    }
  }

  return [pscustomobject]@{
    path = $Path
    name = $profileName
    displayName = [string](Get-OptionalPropertyValue -InputObject $Profile -Name "displayName")
    metadataFieldCount = @($metadataFields).Count
    sectionFieldCount = @((Get-OptionalPropertyValue -InputObject $Profile -Name "sectionFields")).Count
  }
}

function New-SummaryObject {
  param(
    [AllowEmptyCollection()]
    [object[]]$Profiles,

    [AllowEmptyCollection()]
    [object[]]$Findings,

    [Parameter(Mandatory = $true)]
    [string]$SchemaPath
  )

  $findingCountsByCode = @{}
  foreach ($group in ($Findings | Group-Object -Property code)) {
    $findingCountsByCode[[string]$group.Name] = [int]$group.Count
  }

  $errorCount = @($Findings | Where-Object { [string]$_.severity -eq "error" }).Count
  $warningCount = @($Findings | Where-Object { [string]$_.severity -eq "warning" }).Count

  return [pscustomobject]@{
    passed = ($errorCount -eq 0)
    schemaPath = $SchemaPath
    profiles = @($Profiles)
    findings = @($Findings)
    summary = [pscustomobject]@{
      profileCount = @($Profiles).Count
      errorCount = $errorCount
      warningCount = $warningCount
      findingCountsByCode = $findingCountsByCode
    }
  }
}

function Format-ValidationResultText {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Result
  )

  $lines = New-Object System.Collections.Generic.List[string]
  if ([bool]$Result.passed) {
    [void]$lines.Add(("OK: validated {0} report profile(s)." -f [int]$Result.summary.profileCount))
  } else {
    [void]$lines.Add(("FAILED: {0} error(s), {1} warning(s)." -f [int]$Result.summary.errorCount, [int]$Result.summary.warningCount))
  }

  foreach ($finding in @($Result.findings)) {
    [void]$lines.Add(("[{0}] {1} {2} - {3}" -f [string]$finding.severity, [string]$finding.code, [string]$finding.path, [string]$finding.message))
  }

  return ($lines -join [Environment]::NewLine)
}

function Write-ValidationOutput {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Text,

    [AllowNull()]
    [string]$OutputPath
  )

  if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $parent = Split-Path -Parent $OutputPath
    if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent)) {
      New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($OutputPath, $Text, $utf8NoBom)
  }

  Write-Output $Text
}

$findings = New-Object System.Collections.Generic.List[object]
$profilePaths = @(Get-ResolvedProfilePaths -ProfileDirectory $ProfileDir -RequestedProfilePaths $ProfilePath -Findings $findings)
$profileResults = New-Object System.Collections.Generic.List[object]
$profilePathsByName = @{}

foreach ($resolvedProfilePath in $profilePaths) {
  $profile = Read-ProfileJson -Path $resolvedProfilePath -Findings $findings
  if ($null -eq $profile) {
    continue
  }

  $profileResult = Test-ReportProfile -Path $resolvedProfilePath -Profile $profile -Findings $findings
  [void]$profileResults.Add($profileResult)
  if (-not [string]::IsNullOrWhiteSpace([string]$profileResult.name)) {
    $nameKey = ([string]$profileResult.name).ToLowerInvariant()
    if (-not $profilePathsByName.ContainsKey($nameKey)) {
      $profilePathsByName[$nameKey] = New-Object System.Collections.Generic.List[string]
    }
    [void]$profilePathsByName[$nameKey].Add($resolvedProfilePath)
  }
}

foreach ($nameKey in $profilePathsByName.Keys) {
  if ($profilePathsByName[$nameKey].Count -gt 1) {
    foreach ($duplicatePath in $profilePathsByName[$nameKey]) {
      Add-ProfileFinding -Findings $findings -Severity error -Code "profile-name-duplicate" -Path $duplicatePath -Message "Profile name '$nameKey' is used by more than one file."
    }
  }
}

if ([string]::IsNullOrWhiteSpace($ProfileDir)) {
  $repoRootForSchema = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
  $resolvedSchemaPath = Join-Path (Join-Path $repoRootForSchema "profiles") "report-profile.schema.json"
} else {
  $resolvedSchemaPath = Join-Path $ProfileDir "report-profile.schema.json"
}

$result = New-SummaryObject -Profiles $profileResults.ToArray() -Findings $findings.ToArray() -SchemaPath $resolvedSchemaPath
if ($Format -eq "json") {
  $outputText = $result | ConvertTo-Json -Depth 12
} else {
  $outputText = Format-ValidationResultText -Result $result
}

Write-ValidationOutput -Text $outputText -OutputPath $OutFile

if (-not [bool]$result.passed) {
  exit 1
}
