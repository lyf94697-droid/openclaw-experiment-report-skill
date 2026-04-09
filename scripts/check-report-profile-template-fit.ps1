[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$TemplatePath,

  [string]$ReportPath,

  [string]$ReportText,

  [string]$MetadataPath,

  [string]$MetadataJson,

  [string]$ReportProfileName = "experiment-report",

  [string]$ReportProfilePath,

  [ValidateSet("json", "markdown")]
  [string]$Format = "json",

  [string]$OutFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "report-profiles.ps1")

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$reportProfile = Get-ReportProfile -ProfileName $ReportProfileName -ProfilePath $ReportProfilePath -RepoRoot $repoRoot
$sectionRules = @(Get-ReportProfileSectionRules -Profile $reportProfile)
$metadataPrefixes = @(Get-ReportProfileMetadataPrefixes -Profile $reportProfile)

function Get-ObjectPropertyValue {
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

function Get-StringArray {
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
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
    )
  }

  if ([string]::IsNullOrWhiteSpace([string]$Value)) {
    return @()
  }

  return @([string]$Value)
}

function Get-DiagnosticContextValue {
  param(
    [AllowNull()]
    [object]$Diagnostic,

    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  $context = Get-ObjectPropertyValue -Object $Diagnostic -Name "context"
  if ($null -eq $context) {
    return $null
  }

  return Get-ObjectPropertyValue -Object $context -Name $Name
}

function Add-GroupedOccurrence {
  param(
    [Parameter(Mandatory = $true)]
    [hashtable]$Table,

    [Parameter(Mandatory = $true)]
    [string]$Key,

    [Parameter(Mandatory = $true)]
    [object]$Diagnostic
  )

  if ([string]::IsNullOrWhiteSpace($Key)) {
    return
  }

  if (-not $Table.ContainsKey($Key)) {
    $Table[$Key] = [ordered]@{
      key = $Key
      count = 0
      locations = New-Object System.Collections.Generic.List[string]
      samples = New-Object System.Collections.Generic.List[object]
    }
  }

  $entry = $Table[$Key]
  $entry.count++

  $location = [string](Get-DiagnosticContextValue -Diagnostic $Diagnostic -Name "location")
  if (-not [string]::IsNullOrWhiteSpace($location) -and -not $entry.locations.Contains($location)) {
    $entry.locations.Add($location) | Out-Null
  }

  if ($entry.samples.Count -lt 3) {
    $entry.samples.Add($Diagnostic) | Out-Null
  }
}

function Get-SectionRuleById {
  param(
    [AllowNull()]
    [string]$SectionId
  )

  foreach ($rule in $sectionRules) {
    if ([string]$rule.id -eq [string]$SectionId) {
      return $rule
    }
  }

  return $null
}

function Get-SectionHeuristicKeywords {
  param(
    [AllowNull()]
    [string]$SectionId
  )

  switch ([string]$SectionId) {
    "purpose" { return @('\u76EE\u7684', '\u76EE\u6807') | ForEach-Object { [regex]::Unescape($_) } }
    "environment" { return @('\u73AF\u5883', '\u5668\u6750', '\u8BBE\u5907', '\u62D3\u6251', '\u88C5\u7F6E', '\u5E73\u53F0') | ForEach-Object { [regex]::Unescape($_) } }
    "theory" { return @('\u539F\u7406', '\u8981\u6C42', '\u4EFB\u52A1', '\u7406\u8BBA') | ForEach-Object { [regex]::Unescape($_) } }
    "steps" { return @('\u6B65\u9AA4', '\u8FC7\u7A0B', '\u64CD\u4F5C') | ForEach-Object { [regex]::Unescape($_) } }
    "result" { return @('\u7ED3\u679C', '\u73B0\u8C61', '\u8BB0\u5F55') | ForEach-Object { [regex]::Unescape($_) } }
    "analysis" { return @('\u5206\u6790', '\u8BA8\u8BBA', '\u95EE\u9898') | ForEach-Object { [regex]::Unescape($_) } }
    "summary" { return @('\u603B\u7ED3', '\u5C0F\u7ED3', '\u601D\u8003') | ForEach-Object { [regex]::Unescape($_) } }
    default { return @() }
  }
}

function Resolve-SuggestedSectionRule {
  param(
    [AllowNull()]
    [string]$Heading
  )

  if ([string]::IsNullOrWhiteSpace($Heading)) {
    return $null
  }

  $bestRule = $null
  $bestScore = 0
  foreach ($rule in $sectionRules) {
    $score = 0
    foreach ($keyword in @(Get-SectionHeuristicKeywords -SectionId ([string]$rule.id))) {
      if (-not [string]::IsNullOrWhiteSpace($keyword) -and $Heading -match [regex]::Escape($keyword)) {
        $score += 3
      }
    }

    foreach ($alias in @($rule.headingAliases)) {
      $aliasText = [string]$alias
      if (-not [string]::IsNullOrWhiteSpace($aliasText) -and $Heading -match [regex]::Escape($aliasText)) {
        $score += 1
      }
    }

    if ($score -gt $bestScore) {
      $bestScore = $score
      $bestRule = $rule
    }
  }

  if ($bestScore -le 0) {
    return $null
  }

  return $bestRule
}

function Split-CompositeMatchTokens {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return @()
  }

  $parts = New-Object System.Collections.Generic.List[string]
  $splitPattern = [regex]::Unescape('(?:/|\\|\||,|\uFF0C|;|\uFF1B|\r?\n)+')
  foreach ($rawPart in ($Text -split $splitPattern)) {
    $trimmed = (($rawPart -replace '\s+', ' ').Trim())
    if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
      $parts.Add($trimmed) | Out-Null
    }
  }

  if ($parts.Count -eq 0) {
    return @([string]$Text)
  }

  return @($parts | Select-Object -Unique)
}

function ConvertTo-MarkdownCode {
  param(
    [AllowNull()]
    [object]$Value
  )

  if ($null -eq $Value) {
    return ""
  }

  return (($Value | ConvertTo-Json -Depth 8) -join [Environment]::NewLine)
}

$fieldMapJson = & (Join-Path $PSScriptRoot "generate-docx-field-map.ps1") `
  -TemplatePath $TemplatePath `
  -ReportPath $ReportPath `
  -ReportText $ReportText `
  -MetadataPath $MetadataPath `
  -MetadataJson $MetadataJson `
  -ReportProfileName ([string]$reportProfile.name) `
  -ReportProfilePath ([string]$reportProfile.resolvedProfilePath) `
  -Format json | Out-String

$fieldMapResult = $fieldMapJson | ConvertFrom-Json
if ($null -eq $fieldMapResult) {
  throw "Failed to parse generate-docx-field-map output."
}

$diagnostics = @(Get-ObjectPropertyValue -Object $fieldMapResult -Name "diagnostics")

$missingMetadataTable = @{}
$unrecognizedMetadataTable = @{}
$missingSectionTable = @{}
$unrecognizedSectionTable = @{}
$unmatchedCompositeTable = @{}

foreach ($diagnostic in $diagnostics) {
  $code = [string](Get-ObjectPropertyValue -Object $diagnostic -Name "code")
  switch ($code) {
    "missing_metadata_value" {
      Add-GroupedOccurrence -Table $missingMetadataTable -Key ([string](Get-DiagnosticContextValue -Diagnostic $diagnostic -Name "label")) -Diagnostic $diagnostic
    }
    "unrecognized_template_metadata_label" {
      Add-GroupedOccurrence -Table $unrecognizedMetadataTable -Key ([string](Get-DiagnosticContextValue -Diagnostic $diagnostic -Name "label")) -Diagnostic $diagnostic
    }
    "missing_report_section" {
      Add-GroupedOccurrence -Table $missingSectionTable -Key ([string](Get-DiagnosticContextValue -Diagnostic $diagnostic -Name "heading")) -Diagnostic $diagnostic
    }
    "unrecognized_template_section_heading" {
      Add-GroupedOccurrence -Table $unrecognizedSectionTable -Key ([string](Get-DiagnosticContextValue -Diagnostic $diagnostic -Name "heading")) -Diagnostic $diagnostic
    }
    "unmatched_composite_template_cell" {
      Add-GroupedOccurrence -Table $unmatchedCompositeTable -Key ([string](Get-DiagnosticContextValue -Diagnostic $diagnostic -Name "cellText")) -Diagnostic $diagnostic
    }
  }
}

$metadataFieldsToAdd = New-Object System.Collections.Generic.List[object]
foreach ($entry in @($unrecognizedMetadataTable.Values | Sort-Object key)) {
  $label = [string]$entry.key
  $metadataFieldsToAdd.Add([pscustomobject][ordered]@{
      label = $label
      count = [int]$entry.count
      locations = @($entry.locations)
      knownMetadataLabels = @($metadataPrefixes)
      suggestedProfilePatch = [ordered]@{
        key = "TODO_MetadataKey"
        label = $label
      }
    }) | Out-Null
}

$missingMetadataValues = New-Object System.Collections.Generic.List[object]
foreach ($entry in @($missingMetadataTable.Values | Sort-Object key)) {
  $firstSample = if ($entry.samples.Count -gt 0) { $entry.samples[0] } else { $null }
  $metadataId = [string](Get-DiagnosticContextValue -Diagnostic $firstSample -Name "metadataId")
  $missingMetadataValues.Add([pscustomobject][ordered]@{
      label = [string]$entry.key
      metadataId = $metadataId
      count = [int]$entry.count
      locations = @($entry.locations)
    }) | Out-Null
}

$sectionAliasesToAdd = New-Object System.Collections.Generic.List[object]
foreach ($entry in @($unrecognizedSectionTable.Values | Sort-Object key)) {
  $heading = [string]$entry.key
  $suggestedRule = Resolve-SuggestedSectionRule -Heading $heading
  if ($null -ne $suggestedRule) {
    $sectionAliasesToAdd.Add([pscustomobject][ordered]@{
        heading = $heading
        count = [int]$entry.count
        locations = @($entry.locations)
        suggestedSectionId = [string]$suggestedRule.id
        suggestedCanonicalLabel = [string]$suggestedRule.canonicalLabel
        suggestedProfilePatch = [ordered]@{
          sectionId = [string]$suggestedRule.id
          addAliases = @($heading)
        }
      }) | Out-Null
  } else {
    $sectionAliasesToAdd.Add([pscustomobject][ordered]@{
        heading = $heading
        count = [int]$entry.count
        locations = @($entry.locations)
        suggestedSectionId = $null
        suggestedCanonicalLabel = $null
        suggestedProfilePatch = [ordered]@{
          key = "TODO_SectionKey"
          heading = $heading
          aliases = @($heading)
        }
      }) | Out-Null
  }
}

$missingReportSections = New-Object System.Collections.Generic.List[object]
foreach ($entry in @($missingSectionTable.Values | Sort-Object key)) {
  $firstSample = if ($entry.samples.Count -gt 0) { $entry.samples[0] } else { $null }
  $sectionId = [string](Get-DiagnosticContextValue -Diagnostic $firstSample -Name "sectionId")
  $sectionRule = Get-SectionRuleById -SectionId $sectionId
  $missingReportSections.Add([pscustomobject][ordered]@{
      heading = [string]$entry.key
      sectionId = $sectionId
      canonicalLabel = if ($null -ne $sectionRule) { [string]$sectionRule.canonicalLabel } else { "" }
      count = [int]$entry.count
      locations = @($entry.locations)
    }) | Out-Null
}

$compositeRulesToAdd = New-Object System.Collections.Generic.List[object]
foreach ($entry in @($unmatchedCompositeTable.Values | Sort-Object key)) {
  $firstSample = if ($entry.samples.Count -gt 0) { $entry.samples[0] } else { $null }
  $matchedSectionIds = @(
    Get-StringArray -Value (Get-DiagnosticContextValue -Diagnostic $firstSample -Name "matchedSectionIds") |
      Select-Object -Unique
  )
  $blocks = New-Object System.Collections.Generic.List[object]
  foreach ($sectionId in $matchedSectionIds) {
    $sectionRule = Get-SectionRuleById -SectionId $sectionId
    $heading = if ($null -ne $sectionRule -and -not [string]::IsNullOrWhiteSpace([string]$sectionRule.canonicalLabel)) {
      [string]$sectionRule.canonicalLabel
    } else {
      [string]$sectionId
    }
    $blocks.Add([ordered]@{
        heading = $heading
        sectionIds = @([string]$sectionId)
      }) | Out-Null
  }

  $compositeRulesToAdd.Add([pscustomobject][ordered]@{
      cellText = [string]$entry.key
      count = [int]$entry.count
      locations = @($entry.locations)
      matchedSectionIds = @($matchedSectionIds)
      suggestedProfilePatch = [ordered]@{
        matchAll = @(Split-CompositeMatchTokens -Text ([string]$entry.key))
        blocks = [object[]]$blocks
      }
    }) | Out-Null
}

$warningCount = 0
$errorCount = 0
$infoCount = 0
foreach ($diagnostic in $diagnostics) {
  switch ([string](Get-ObjectPropertyValue -Object $diagnostic -Name "severity")) {
    "error" { $errorCount++ }
    "info" { $infoCount++ }
    default { $warningCount++ }
  }
}

$profileChangeSuggestionCount = $metadataFieldsToAdd.Count + $sectionAliasesToAdd.Count + $compositeRulesToAdd.Count
$inputGapCount = $missingMetadataValues.Count + $missingReportSections.Count

$result = [ordered]@{
  templatePath = [string](Get-ObjectPropertyValue -Object $fieldMapResult -Name "templatePath")
  reportSource = [string](Get-ObjectPropertyValue -Object $fieldMapResult -Name "reportSource")
  reportProfileName = [string](Get-ObjectPropertyValue -Object $fieldMapResult -Name "reportProfileName")
  reportProfilePath = [string](Get-ObjectPropertyValue -Object $fieldMapResult -Name "reportProfilePath")
  reportInputMode = [string](Get-ObjectPropertyValue -Object $fieldMapResult -Name "reportInputMode")
  metadataInputMode = [string](Get-ObjectPropertyValue -Object $fieldMapResult -Name "metadataInputMode")
  summary = [ordered]@{
    diagnosticCount = @($diagnostics).Count
    warningCount = $warningCount
    errorCount = $errorCount
    infoCount = $infoCount
    profileChangeSuggestionCount = $profileChangeSuggestionCount
    inputGapCount = $inputGapCount
    readyForNewProfile = ($profileChangeSuggestionCount -eq 0)
  }
  suggestions = [ordered]@{
    metadataFieldsToAdd = ([object[]]$metadataFieldsToAdd)
    sectionAliasesToAdd = ([object[]]$sectionAliasesToAdd)
    compositeRulesToAdd = ([object[]]$compositeRulesToAdd)
  }
  inputGaps = [ordered]@{
    missingMetadataValues = ([object[]]$missingMetadataValues)
    missingReportSections = ([object[]]$missingReportSections)
  }
  fieldMapSummary = [ordered]@{
    fieldCount = [int](Get-ObjectPropertyValue -Object (Get-ObjectPropertyValue -Object $fieldMapResult -Name "summary") -Name "fieldCount")
    mappedMetadataCount = [int](Get-ObjectPropertyValue -Object (Get-ObjectPropertyValue -Object $fieldMapResult -Name "summary") -Name "mappedMetadataCount")
    mappedSectionCount = [int](Get-ObjectPropertyValue -Object (Get-ObjectPropertyValue -Object $fieldMapResult -Name "summary") -Name "mappedSectionCount")
    diagnosticCountsByCode = (Get-ObjectPropertyValue -Object (Get-ObjectPropertyValue -Object $fieldMapResult -Name "summary") -Name "diagnosticCountsByCode")
  }
  diagnostics = ([object[]]$diagnostics)
}

if ($Format -eq "json") {
  $output = $result | ConvertTo-Json -Depth 10
} else {
  $lines = New-Object System.Collections.Generic.List[string]
  [void]$lines.Add("# Report Profile Template Fit")
  [void]$lines.Add("")
  [void]$lines.Add("- Template: $($result.templatePath)")
  [void]$lines.Add("- Report source: $($result.reportSource)")
  [void]$lines.Add("- Report profile: $($result.reportProfileName)")
  [void]$lines.Add("- Diagnostics: $($result.summary.diagnosticCount)")
  [void]$lines.Add("- Profile changes suggested: $($result.summary.profileChangeSuggestionCount)")
  [void]$lines.Add("- Input gaps: $($result.summary.inputGapCount)")
  [void]$lines.Add("")

  [void]$lines.Add("## Profile Changes")
  if ($result.summary.profileChangeSuggestionCount -eq 0) {
    [void]$lines.Add("- None")
  } else {
    foreach ($item in @($result.suggestions.metadataFieldsToAdd)) {
      [void]$lines.Add("- Add metadata field or alias: $($item.label) at $((@($item.locations) -join ', '))")
    }
    foreach ($item in @($result.suggestions.sectionAliasesToAdd)) {
      $targetText = if (-not [string]::IsNullOrWhiteSpace([string]$item.suggestedSectionId)) {
        " -> suggest section '$([string]$item.suggestedSectionId)'"
      } else {
        ""
      }
      [void]$lines.Add("- Add section alias or section field: $($item.heading)$targetText")
    }
    foreach ($item in @($result.suggestions.compositeRulesToAdd)) {
      [void]$lines.Add("- Add composite rule for cell: $($item.cellText)")
    }
  }
  [void]$lines.Add("")

  [void]$lines.Add("## Input Gaps")
  if ($result.summary.inputGapCount -eq 0) {
    [void]$lines.Add("- None")
  } else {
    foreach ($item in @($result.inputGaps.missingMetadataValues)) {
      [void]$lines.Add("- Missing metadata value: $($item.label)")
    }
    foreach ($item in @($result.inputGaps.missingReportSections)) {
      [void]$lines.Add("- Missing report section content: $($item.heading)")
    }
  }
  [void]$lines.Add("")

  [void]$lines.Add("## Suggested Profile Patches")
  if ($result.summary.profileChangeSuggestionCount -eq 0) {
    [void]$lines.Add("- None")
  } else {
    foreach ($item in @($result.suggestions.metadataFieldsToAdd)) {
      [void]$lines.Add("- Metadata label: $($item.label)")
      [void]$lines.Add('```json')
      [void]$lines.Add((ConvertTo-MarkdownCode -Value $item.suggestedProfilePatch))
      [void]$lines.Add('```')
    }
    foreach ($item in @($result.suggestions.sectionAliasesToAdd)) {
      [void]$lines.Add("- Section heading: $($item.heading)")
      [void]$lines.Add('```json')
      [void]$lines.Add((ConvertTo-MarkdownCode -Value $item.suggestedProfilePatch))
      [void]$lines.Add('```')
    }
    foreach ($item in @($result.suggestions.compositeRulesToAdd)) {
      [void]$lines.Add("- Composite cell: $($item.cellText)")
      [void]$lines.Add('```json')
      [void]$lines.Add((ConvertTo-MarkdownCode -Value $item.suggestedProfilePatch))
      [void]$lines.Add('```')
    }
  }
  [void]$lines.Add("")

  [void]$lines.Add("## Diagnostics")
  if (@($result.diagnostics).Count -eq 0) {
    [void]$lines.Add("- None")
  } else {
    foreach ($diagnostic in @($result.diagnostics)) {
      $line = "- [$([string]$diagnostic.severity)] [$([string]$diagnostic.code)] $([string]$diagnostic.message)"
      if ($diagnostic.PSObject.Properties.Name -contains "suggestion" -and -not [string]::IsNullOrWhiteSpace([string]$diagnostic.suggestion)) {
        $line = "$line Suggestion: $([string]$diagnostic.suggestion)"
      }
      [void]$lines.Add($line)
    }
  }

  $output = $lines -join [Environment]::NewLine
}

if ([string]::IsNullOrWhiteSpace($OutFile)) {
  Write-Output $output
} else {
  $directory = Split-Path -Parent $OutFile
  if (-not [string]::IsNullOrWhiteSpace($directory)) {
    [System.IO.Directory]::CreateDirectory($directory) | Out-Null
  }
  [System.IO.File]::WriteAllText($OutFile, $output, (New-Object System.Text.UTF8Encoding($true)))
  Write-Output "Wrote template-fit report to $OutFile"
}
