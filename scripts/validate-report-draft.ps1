[CmdletBinding()]
param(
  [string]$Path,

  [string]$Text,

  [string]$ReportProfileName = "experiment-report",

  [string]$ReportProfilePath,

  [string]$RequirementsPath,

  [string]$RequirementsJson,

  [ValidateSet("markdown", "json")]
  [string]$Format = "markdown",

  [string]$OutFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "report-profiles.ps1")

function ConvertTo-PlainHashtable {
  param(
    [Parameter(Mandatory = $true)]
    [object]$InputObject
  )

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

function Get-StringList {
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
      foreach ($text in (Get-StringList -Value $item)) {
        if (-not [string]::IsNullOrWhiteSpace($text)) {
          [void]$items.Add($text)
        }
      }
    }
    return $items.ToArray()
  }

  $textValue = [string]$Value
  if (-not [string]::IsNullOrWhiteSpace($textValue)) {
    [void]$items.Add($textValue.Trim())
  }

  return $items.ToArray()
}

function Get-ReportText {
  param(
    [AllowNull()]
    [string]$TextPath,

    [AllowNull()]
    [string]$InlineText
  )

  if ([string]::IsNullOrWhiteSpace($TextPath) -eq [string]::IsNullOrWhiteSpace($InlineText)) {
    throw "Provide exactly one of -Path or -Text."
  }

  if (-not [string]::IsNullOrWhiteSpace($TextPath)) {
    $resolvedPath = Resolve-Path -LiteralPath $TextPath
    return @{
      Source = $resolvedPath.Path
      Text = Get-Content -LiteralPath $resolvedPath.Path -Raw -Encoding UTF8
    }
  }

  return @{
    Source = "[inline text]"
    Text = $InlineText
  }
}

function Get-RequirementsRoot {
  param(
    [AllowNull()]
    [string]$ConfigPath,

    [AllowNull()]
    [string]$ConfigJson,

    [AllowNull()]
    [string]$ProfileName = "experiment-report",

    [AllowNull()]
    [string]$ProfilePath,

    [AllowNull()]
    [string]$RepoRoot
  )

  if ([string]::IsNullOrWhiteSpace($ConfigPath) -and [string]::IsNullOrWhiteSpace($ConfigJson)) {
    $profile = Get-ReportProfile -ProfileName $ProfileName -ProfilePath $ProfilePath -RepoRoot $RepoRoot
    $detailProfile = Get-ReportProfileDetailProfile -Profile $profile -DetailLevel "standard"
    $sections = foreach ($sectionField in (Get-ReportProfileSectionFields -Profile $profile)) {
      $minChars = 0
      if ($null -ne $sectionField.minChars -and $sectionField.minChars.PSObject.Properties.Name -contains "standard") {
        $minChars = [int]$sectionField.minChars.standard
      }

      [pscustomobject]@{
        name = [string]$sectionField.heading
        aliases = @($sectionField.aliases)
        minChars = $minChars
      }
    }

    return @{
      courseName = ""
      experimentName = ""
      minChars = [int]$detailProfile.minChars
      sections = @($sections)
      forbiddenPatterns = @($profile.forbiddenPatterns)
      paginationRiskThresholds = Get-ReportProfilePaginationRiskThresholds -Profile $profile
      reportProfileName = [string]$profile.name
      reportProfilePath = [string]$profile.resolvedProfilePath
    }
  } elseif ([string]::IsNullOrWhiteSpace($ConfigPath) -eq [string]::IsNullOrWhiteSpace($ConfigJson)) {
    throw "Provide zero or one of -RequirementsPath and -RequirementsJson."
  }

  if (-not [string]::IsNullOrWhiteSpace($ConfigPath)) {
    $resolvedPath = Resolve-Path -LiteralPath $ConfigPath
    return ConvertTo-PlainHashtable -InputObject ((Get-Content -LiteralPath $resolvedPath.Path -Raw -Encoding UTF8) | ConvertFrom-Json)
  }

  return ConvertTo-PlainHashtable -InputObject ($ConfigJson | ConvertFrom-Json)
}

function Resolve-PaginationRiskThresholds {
  param(
    [AllowNull()]
    [object]$Value
  )

  $thresholds = [ordered]@{
    longSectionChars = 900
    denseSectionChars = 550
    denseSectionParagraphs = 2
    figureClusterRefs = 3
  }

  if ($null -eq $Value) {
    return [pscustomobject]$thresholds
  }

  $configuredThresholds = ConvertTo-PlainHashtable -InputObject $Value
  foreach ($key in @($thresholds.Keys)) {
    if ($configuredThresholds.ContainsKey($key) -and $null -ne $configuredThresholds[$key] -and -not [string]::IsNullOrWhiteSpace([string]$configuredThresholds[$key])) {
      $thresholds[$key] = [int]$configuredThresholds[$key]
    }
  }

  return [pscustomobject]$thresholds
}

function Test-TextContains {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Haystack,

    [Parameter(Mandatory = $true)]
    [string]$Needle
  )

  return $Haystack.ToLowerInvariant().Contains($Needle.ToLowerInvariant())
}

function Get-HeadingPattern {
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$Aliases
  )

  $escapedAliases = foreach ($alias in $Aliases) {
    [regex]::Escape($alias)
  }

  return '^(?:\u7B2C?[0-9\u4E00\u4E8C\u4E09\u56DB\u4E94\u516D\u4E03\u516B\u4E5D\u5341]+(?:\u7AE0|\u8282)?[.\)\u3001]?\s*)?(?:' + ($escapedAliases -join '|') + ')\s*(?:[:\uFF1A])?$'
}

function Get-ValidationRemediation {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Code
  )

  switch ($Code) {
    "missing-profile-required-heading" { return "Add the missing profile section heading and body, or add an alias to the active report profile if the report uses a valid alternate heading." }
    "missing-required-section" { return "Add the missing required section heading and body, or remove the section from the requirements if it is not required for this run." }
    "duplicate-section-heading" { return "Keep one canonical heading and merge or rename repeated sections so downstream template filling has a single target." }
    "section-order-anomaly" { return "Reorder the sections to match the active profile or requirements order, or update the profile sectionFields order if the template expects this sequence." }
    "empty-section" { return "Write concrete body content under the section, or remove that section from the active profile or requirements if it is not required." }
    "placeholder-only-section" { return "Replace placeholders such as TODO, placeholder text, or blank underline blocks with final report content before building the docx." }
    "short-section" { return "Expand the section with concrete steps, evidence, results, or analysis until it meets minChars, or tune the profile threshold if the section is intentionally brief." }
    "pagination-risk-long-section" { return "Split the long section into shorter paragraphs or subsections, or raise paginationRiskThresholds.longSectionChars in the active profile if this length is expected." }
    "pagination-risk-dense-section-block" { return "Break the dense text block into more paragraphs or list items, or tune denseSectionChars and denseSectionParagraphs in the active profile." }
    "pagination-risk-figure-cluster" { return "Move some figure references to adjacent sections, group screenshots deliberately, or raise paginationRiskThresholds.figureClusterRefs if this density is normal." }
    "missing-course-name" { return "Add the required course name near the top of the report, or fix courseName in the requirements if the expected value is wrong." }
    "missing-experiment-name" { return "Add the required experiment or project name near the top of the report, or fix experimentName in the requirements if the expected value is wrong." }
    "missing-keyword" { return "Add the required keyword in the relevant section, or remove it from requiredKeywords if it is not required for this document." }
    "missing-required-phrase" { return "Add the required phrase in the relevant section, or remove it from requiredPhrases if it is no longer required." }
    "forbidden-pattern" { return "Remove placeholder or generated boilerplate wording before submission." }
    "short-report" { return "Expand the report with concrete procedure, result, and analysis content, or tune the detail profile minChars if this document type is intentionally shorter." }
    "missing-figure-refs" { return "Add figure references and captions for the required screenshots, or lower minFigureRefs if screenshots are not required for this run." }
    default { return $null }
  }
}

function Add-Finding {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Target,

    [Parameter(Mandatory = $true)]
    [string]$Severity,

    [Parameter(Mandatory = $true)]
    [string]$Code,

    [Parameter(Mandatory = $true)]
    [string]$Message,

    [AllowNull()]
    [string]$Category,

    [AllowNull()]
    [object]$Context,

    [AllowNull()]
    [string]$Remediation
  )

  $contextObject = $null
  if ($null -ne $Context) {
    $contextTable = ConvertTo-PlainHashtable -InputObject $Context
    if ($contextTable.Count -gt 0) {
      $contextObject = [pscustomobject]$contextTable
    }
  }

  $resolvedRemediation = if ([string]::IsNullOrWhiteSpace($Remediation)) {
    Get-ValidationRemediation -Code $Code
  } else {
    $Remediation
  }

  $Target.Add([pscustomobject]@{
      severity = $Severity
      code = $Code
      category = $Category
      message = $Message
      remediation = $resolvedRemediation
      context = $contextObject
    }) | Out-Null
}

function New-SectionRule {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Rule
  )

  $table = ConvertTo-PlainHashtable -InputObject $Rule
  if (-not $table.ContainsKey("name")) {
    throw "Section rules must include a name."
  }

  $name = [string]$table["name"]
  $aliases = @(Get-StringList -Value $table["aliases"])
  if ($aliases.Count -eq 0) {
    $aliases = @($name)
  }

  $minChars = 0
  if ($table.ContainsKey("minChars") -and -not [string]::IsNullOrWhiteSpace([string]$table["minChars"])) {
    $minChars = [int]$table["minChars"]
  }

  return [pscustomobject]@{
    name = $name
    aliases = $aliases
    minChars = $minChars
    headingPattern = Get-HeadingPattern -Aliases $aliases
  }
}

function Normalize-ContentChars {
  param(
    [Parameter(Mandatory = $true)]
    [AllowEmptyString()]
    [string]$Text
  )

  return (($Text -replace '\s+', '')).Length
}

function Test-PlaceholderLine {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $false
  }

  $trimmed = $Text.Trim()
  $collapsed = $trimmed -replace '\s+', ''
  if ([string]::IsNullOrWhiteSpace($collapsed)) {
    return $false
  }

  if ($collapsed -match '^[\-_=\.\*~#]{3,}$') {
    return $true
  }

  if ($collapsed -match '^[xX]{3,}$') {
    return $true
  }

  $normalizedToken = $collapsed.Trim([char[]]@('(', ')', '[', ']')).ToLowerInvariant()
  if (@(
      'todo',
      'tbd',
      'placeholder',
      'fill-me',
      'replace-me',
      'dummy',
      'n/a',
      'na'
    ) -contains $normalizedToken) {
    return $true
  }

  if ($trimmed.Length -le 40 -and $trimmed -match 'placeholder|TODO|TBD|fill-me|replace-me|dummy|insert here|to be filled') {
    return $true
  }

  return $false
}

function Get-SectionContentMetrics {
  param(
    [AllowEmptyCollection()]
    [string[]]$ContentLines
  )

  $contentText = (@($ContentLines) -join [Environment]::NewLine).Trim()
  $trimmedContentLines = @(
    foreach ($contentLine in @($ContentLines)) {
      $trimmedLine = [string]$contentLine
      if (-not [string]::IsNullOrWhiteSpace($trimmedLine)) {
        $trimmedLine.Trim()
      }
    }
  )

  $placeholderLineCount = @(
    foreach ($trimmedContentLine in $trimmedContentLines) {
      if (Test-PlaceholderLine -Text $trimmedContentLine) {
        $trimmedContentLine
      }
    }
  ).Count

  return [pscustomobject]@{
    contentChars = Normalize-ContentChars -Text $contentText
    nonEmptyLineCount = $trimmedContentLines.Count
    paragraphCount = $trimmedContentLines.Count
    figureRefCount = @([regex]::Matches($contentText, '\u56FE\s*\d+')).Count
    empty = ($trimmedContentLines.Count -eq 0)
    placeholderOnly = (($trimmedContentLines.Count -gt 0) -and ($placeholderLineCount -eq $trimmedContentLines.Count))
  }
}

function Get-FindingCountTable {
  param(
    [AllowNull()]
    [object]$Items,

    [Parameter(Mandatory = $true)]
    [string]$PropertyName
  )

  $table = @{}
  $sourceItems = New-Object System.Collections.Generic.List[object]
  if ($null -ne $Items) {
    if (($Items -is [System.Collections.IEnumerable]) -and ($Items -isnot [string])) {
      foreach ($entry in $Items) {
        $sourceItems.Add($entry) | Out-Null
      }
    } else {
      $sourceItems.Add($Items) | Out-Null
    }
  }

  foreach ($item in $sourceItems) {
    if ($null -eq $item -or -not ($item.PSObject.Properties.Name -contains $PropertyName)) {
      continue
    }

    $key = [string]$item.$PropertyName
    if ([string]::IsNullOrWhiteSpace($key)) {
      continue
    }

    if (-not $table.ContainsKey($key)) {
      $table[$key] = 0
    }

    $table[$key] = [int]$table[$key] + 1
  }

  return [pscustomobject]$table
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$usesExternalRequirements = (-not [string]::IsNullOrWhiteSpace($RequirementsPath)) -or (-not [string]::IsNullOrWhiteSpace($RequirementsJson))
$reportInfo = Get-ReportText -TextPath $Path -InlineText $Text
$reportText = [string]$reportInfo.Text
$requirementsRoot = Get-RequirementsRoot -ConfigPath $RequirementsPath -ConfigJson $RequirementsJson -ProfileName $ReportProfileName -ProfilePath $ReportProfilePath -RepoRoot $repoRoot
$paginationRiskThresholdSource = $null
if ($requirementsRoot.ContainsKey("paginationRiskThresholds")) {
  $paginationRiskThresholdSource = $requirementsRoot["paginationRiskThresholds"]
}
$paginationRiskThresholds = Resolve-PaginationRiskThresholds -Value $paginationRiskThresholdSource

$resolvedValidationProfileName = if ($requirementsRoot.ContainsKey("reportProfileName") -and -not [string]::IsNullOrWhiteSpace([string]$requirementsRoot["reportProfileName"])) {
  [string]$requirementsRoot["reportProfileName"]
} elseif (-not $usesExternalRequirements -and -not [string]::IsNullOrWhiteSpace($ReportProfileName)) {
  $ReportProfileName
} else {
  $null
}

$resolvedValidationProfilePath = if ($requirementsRoot.ContainsKey("reportProfilePath") -and -not [string]::IsNullOrWhiteSpace([string]$requirementsRoot["reportProfilePath"])) {
  [string]$requirementsRoot["reportProfilePath"]
} elseif (-not $usesExternalRequirements -and -not [string]::IsNullOrWhiteSpace($ReportProfilePath)) {
  (Resolve-Path -LiteralPath $ReportProfilePath).Path
} else {
  $null
}

$profileBackedValidation = (-not [string]::IsNullOrWhiteSpace($resolvedValidationProfileName)) -and ((-not $usesExternalRequirements) -or $requirementsRoot.ContainsKey("reportProfileName"))

$findings = New-Object System.Collections.Generic.List[object]
$lines = @($reportText -split "\r?\n")
$trimmedLines = @(foreach ($line in $lines) { $line.Trim() })
$normalizedReportChars = Normalize-ContentChars -Text $reportText

$sectionRules = New-Object System.Collections.Generic.List[object]
if ($requirementsRoot.ContainsKey("sections")) {
  $sectionRuleIndex = 0
  foreach ($sectionRule in $requirementsRoot["sections"]) {
    $rule = New-SectionRule -Rule $sectionRule
    $rule | Add-Member -NotePropertyName "ruleIndex" -NotePropertyValue $sectionRuleIndex -Force
    $sectionRules.Add($rule) | Out-Null
    $sectionRuleIndex++
  }
}

$sectionHeadingOccurrences = New-Object System.Collections.Generic.List[object]
for ($lineIndex = 0; $lineIndex -lt $trimmedLines.Count; $lineIndex++) {
  $lineText = $trimmedLines[$lineIndex]
  if ([string]::IsNullOrWhiteSpace($lineText)) {
    continue
  }

  $matchedRule = $null
  foreach ($sectionRule in $sectionRules) {
    if ($lineText -match $sectionRule.headingPattern) {
      $matchedRule = $sectionRule
      break
    }
  }

  if ($null -eq $matchedRule) {
    continue
  }

  $sectionHeadingOccurrences.Add([pscustomobject]@{
      name = $matchedRule.name
      aliases = $matchedRule.aliases
      minChars = $matchedRule.minChars
      ruleIndex = [int]$matchedRule.ruleIndex
      headingIndex = $lineIndex
      lineNumber = $lineIndex + 1
      headingText = $lineText
      contentChars = 0
      nonEmptyLineCount = 0
      paragraphCount = 0
      figureRefCount = 0
      empty = $true
      placeholderOnly = $false
    }) | Out-Null
}

$orderedHeadingOccurrences = @($sectionHeadingOccurrences | Sort-Object headingIndex, ruleIndex)
for ($occurrenceIndex = 0; $occurrenceIndex -lt $orderedHeadingOccurrences.Count; $occurrenceIndex++) {
  $currentOccurrence = $orderedHeadingOccurrences[$occurrenceIndex]
  $nextHeadingIndex = if (($occurrenceIndex + 1) -lt $orderedHeadingOccurrences.Count) { $orderedHeadingOccurrences[$occurrenceIndex + 1].headingIndex } else { $lines.Count }
  $contentLines = if (($currentOccurrence.headingIndex + 1) -lt $nextHeadingIndex) { $lines[($currentOccurrence.headingIndex + 1)..($nextHeadingIndex - 1)] } else { @() }
  $contentMetrics = Get-SectionContentMetrics -ContentLines $contentLines
  $currentOccurrence.contentChars = [int]$contentMetrics.contentChars
  $currentOccurrence.nonEmptyLineCount = [int]$contentMetrics.nonEmptyLineCount
  $currentOccurrence.paragraphCount = [int]$contentMetrics.paragraphCount
  $currentOccurrence.figureRefCount = [int]$contentMetrics.figureRefCount
  $currentOccurrence.empty = [bool]$contentMetrics.empty
  $currentOccurrence.placeholderOnly = [bool]$contentMetrics.placeholderOnly
}

$sectionMatches = New-Object System.Collections.Generic.List[object]
foreach ($sectionRule in $sectionRules) {
  $occurrences = @($orderedHeadingOccurrences | Where-Object { [int]$_.ruleIndex -eq [int]$sectionRule.ruleIndex })
  $primaryOccurrence = if ($occurrences.Count -gt 0) { $occurrences[0] } else { $null }
  $duplicateHeadingLines = if ($occurrences.Count -gt 1) { @($occurrences | Select-Object -Skip 1 -ExpandProperty lineNumber) } else { @() }

  $sectionMatches.Add([pscustomobject]@{
      name = $sectionRule.name
      aliases = $sectionRule.aliases
      minChars = $sectionRule.minChars
      expectedOrder = ([int]$sectionRule.ruleIndex + 1)
      headingIndex = $(if ($null -ne $primaryOccurrence) { $primaryOccurrence.headingIndex } else { $null })
      lineNumber = $(if ($null -ne $primaryOccurrence) { $primaryOccurrence.lineNumber } else { $null })
      headingText = $(if ($null -ne $primaryOccurrence) { $primaryOccurrence.headingText } else { $null })
      occurrenceCount = $occurrences.Count
      duplicateHeadingLines = $duplicateHeadingLines
      contentChars = $(if ($null -ne $primaryOccurrence) { [int]$primaryOccurrence.contentChars } else { 0 })
      nonEmptyLineCount = $(if ($null -ne $primaryOccurrence) { [int]$primaryOccurrence.nonEmptyLineCount } else { 0 })
      paragraphCount = $(if ($null -ne $primaryOccurrence) { [int]$primaryOccurrence.paragraphCount } else { 0 })
      figureRefCount = $(if ($null -ne $primaryOccurrence) { [int]$primaryOccurrence.figureRefCount } else { 0 })
      empty = $(if ($null -ne $primaryOccurrence) { [bool]$primaryOccurrence.empty } else { $false })
      placeholderOnly = $(if ($null -ne $primaryOccurrence) { [bool]$primaryOccurrence.placeholderOnly } else { $false })
      found = ($null -ne $primaryOccurrence)
    }) | Out-Null
}

$missingSectionCode = if ($profileBackedValidation) { "missing-profile-required-heading" } else { "missing-required-section" }

foreach ($sectionMatch in $sectionMatches) {
  if ($sectionMatch.occurrenceCount -gt 1) {
    Add-Finding -Target $findings -Severity "error" -Code "duplicate-section-heading" -Category "structure" -Message ("Section heading '{0}' appears {1} times." -f $sectionMatch.name, $sectionMatch.occurrenceCount) -Context @{
      section = $sectionMatch.name
      lineNumber = $sectionMatch.lineNumber
      duplicateLineNumbers = @($sectionMatch.duplicateHeadingLines)
      occurrenceCount = [int]$sectionMatch.occurrenceCount
    }
  }

  if (-not $sectionMatch.found) {
    Add-Finding -Target $findings -Severity "error" -Code $missingSectionCode -Category "structure" -Message ("Missing required section: {0}" -f $sectionMatch.name) -Context @{
      section = $sectionMatch.name
      expectedOrder = [int]$sectionMatch.expectedOrder
      reportProfileName = $resolvedValidationProfileName
    }
    continue
  }

  if ($sectionMatch.empty) {
    Add-Finding -Target $findings -Severity "error" -Code "empty-section" -Category "structure" -Message ("Section '{0}' has no body content." -f $sectionMatch.name) -Context @{
      section = $sectionMatch.name
      lineNumber = $sectionMatch.lineNumber
    }
    continue
  }

  if ($sectionMatch.placeholderOnly) {
    Add-Finding -Target $findings -Severity "error" -Code "placeholder-only-section" -Category "structure" -Message ("Section '{0}' only contains placeholder content." -f $sectionMatch.name) -Context @{
      section = $sectionMatch.name
      lineNumber = $sectionMatch.lineNumber
      contentChars = [int]$sectionMatch.contentChars
    }
    continue
  }

  if ($sectionMatch.minChars -gt 0 -and $sectionMatch.contentChars -lt $sectionMatch.minChars) {
    Add-Finding -Target $findings -Severity "error" -Code "short-section" -Category "content" -Message ("Section '{0}' is too short: {1} chars < {2}" -f $sectionMatch.name, $sectionMatch.contentChars, $sectionMatch.minChars) -Context @{
      section = $sectionMatch.name
      lineNumber = $sectionMatch.lineNumber
      contentChars = [int]$sectionMatch.contentChars
      minChars = [int]$sectionMatch.minChars
    }
  }

  if ($paginationRiskThresholds.longSectionChars -gt 0 -and $sectionMatch.contentChars -ge $paginationRiskThresholds.longSectionChars) {
    Add-Finding -Target $findings -Severity "warning" -Code "pagination-risk-long-section" -Category "pagination" -Message ("Section '{0}' is long enough to be pagination-sensitive ({1} chars)." -f $sectionMatch.name, $sectionMatch.contentChars) -Context @{
      section = $sectionMatch.name
      lineNumber = $sectionMatch.lineNumber
      contentChars = [int]$sectionMatch.contentChars
      threshold = [int]$paginationRiskThresholds.longSectionChars
    }
  }

  if ($paginationRiskThresholds.denseSectionChars -gt 0 -and $paginationRiskThresholds.denseSectionParagraphs -gt 0 -and $sectionMatch.contentChars -ge $paginationRiskThresholds.denseSectionChars -and $sectionMatch.paragraphCount -le $paginationRiskThresholds.denseSectionParagraphs) {
    Add-Finding -Target $findings -Severity "warning" -Code "pagination-risk-dense-section-block" -Category "pagination" -Message ("Section '{0}' is dense and may split awkwardly across pages ({1} chars in {2} paragraphs)." -f $sectionMatch.name, $sectionMatch.contentChars, $sectionMatch.paragraphCount) -Context @{
      section = $sectionMatch.name
      lineNumber = $sectionMatch.lineNumber
      contentChars = [int]$sectionMatch.contentChars
      paragraphCount = [int]$sectionMatch.paragraphCount
      charThreshold = [int]$paginationRiskThresholds.denseSectionChars
      paragraphThreshold = [int]$paginationRiskThresholds.denseSectionParagraphs
    }
  }

  if ($paginationRiskThresholds.figureClusterRefs -gt 0 -and $sectionMatch.figureRefCount -ge $paginationRiskThresholds.figureClusterRefs) {
    Add-Finding -Target $findings -Severity "warning" -Code "pagination-risk-figure-cluster" -Category "pagination" -Message ("Section '{0}' references many figures ({1}), which may create layout and pagination pressure." -f $sectionMatch.name, $sectionMatch.figureRefCount) -Context @{
      section = $sectionMatch.name
      lineNumber = $sectionMatch.lineNumber
      figureRefCount = [int]$sectionMatch.figureRefCount
      threshold = [int]$paginationRiskThresholds.figureClusterRefs
    }
  }
}

$orderedFoundSections = @($sectionMatches | Where-Object { $_.found } | Sort-Object headingIndex)
for ($matchIndex = 0; $matchIndex -lt ($orderedFoundSections.Count - 1); $matchIndex++) {
  $current = $orderedFoundSections[$matchIndex]
  $next = $orderedFoundSections[$matchIndex + 1]
  if ([int]$current.expectedOrder -gt [int]$next.expectedOrder) {
    Add-Finding -Target $findings -Severity "error" -Code "section-order-anomaly" -Category "structure" -Message ("Section order is inconsistent: '{0}' appears before '{1}', but the expected order is the reverse." -f $current.name, $next.name) -Context @{
      earlierSection = $current.name
      earlierLineNumber = [int]$current.lineNumber
      laterSection = $next.name
      laterLineNumber = [int]$next.lineNumber
      expectedEarlierSection = $next.name
      expectedLaterSection = $current.name
    }
  }
}

$courseName = if ($requirementsRoot.ContainsKey("courseName")) { [string]$requirementsRoot["courseName"] } else { "" }
if (-not [string]::IsNullOrWhiteSpace($courseName) -and -not (Test-TextContains -Haystack $reportText -Needle $courseName)) {
  Add-Finding -Target $findings -Severity "error" -Code "missing-course-name" -Category "content" -Message ("Report does not mention the course name: {0}" -f $courseName)
}

$experimentName = if ($requirementsRoot.ContainsKey("experimentName")) { [string]$requirementsRoot["experimentName"] } else { "" }
if (-not [string]::IsNullOrWhiteSpace($experimentName) -and -not (Test-TextContains -Haystack $reportText -Needle $experimentName)) {
  Add-Finding -Target $findings -Severity "error" -Code "missing-experiment-name" -Category "content" -Message ("Report does not mention the experiment name: {0}" -f $experimentName)
}

foreach ($keyword in (Get-StringList -Value $requirementsRoot["requiredKeywords"])) {
  if (-not (Test-TextContains -Haystack $reportText -Needle $keyword)) {
    Add-Finding -Target $findings -Severity "error" -Code "missing-keyword" -Category "content" -Message ("Report is missing required keyword: {0}" -f $keyword)
  }
}

foreach ($phrase in (Get-StringList -Value $requirementsRoot["requiredPhrases"])) {
  if (-not (Test-TextContains -Haystack $reportText -Needle $phrase)) {
    Add-Finding -Target $findings -Severity "error" -Code "missing-required-phrase" -Category "content" -Message ("Report is missing required phrase: {0}" -f $phrase)
  }
}

foreach ($pattern in (Get-StringList -Value $requirementsRoot["forbiddenPatterns"])) {
  if ([string]::IsNullOrWhiteSpace($pattern)) {
    continue
  }

  if ($reportText -match [regex]::Escape($pattern)) {
    Add-Finding -Target $findings -Severity "error" -Code "forbidden-pattern" -Category "content" -Message ("Report contains forbidden pattern: {0}" -f $pattern)
  }
}

$minChars = 0
if ($requirementsRoot.ContainsKey("minChars") -and -not [string]::IsNullOrWhiteSpace([string]$requirementsRoot["minChars"])) {
  $minChars = [int]$requirementsRoot["minChars"]
}
if ($minChars -gt 0 -and $normalizedReportChars -lt $minChars) {
  Add-Finding -Target $findings -Severity "error" -Code "short-report" -Category "content" -Message ("Report is too short: {0} chars < {1}" -f $normalizedReportChars, $minChars)
}

$minFigureRefs = 0
if ($requirementsRoot.ContainsKey("minFigureRefs") -and -not [string]::IsNullOrWhiteSpace([string]$requirementsRoot["minFigureRefs"])) {
  $minFigureRefs = [int]$requirementsRoot["minFigureRefs"]
}
if ($minFigureRefs -gt 0) {
  $figureRefCount = @([regex]::Matches($reportText, '\u56FE\s*\d+')).Count
  if ($figureRefCount -lt $minFigureRefs) {
    Add-Finding -Target $findings -Severity "error" -Code "missing-figure-refs" -Category "content" -Message ("Report references too few figures: {0} < {1}" -f $figureRefCount, $minFigureRefs)
  }
}

$errorCount = @($findings | Where-Object { $_.severity -eq "error" }).Count
$warningCount = @($findings | Where-Object { $_.severity -eq "warning" }).Count
$paginationRiskCount = @($findings | Where-Object { $_.category -eq "pagination" }).Count
$structuralIssueCount = @($findings | Where-Object { $_.category -eq "structure" }).Count
$findingCountsByCode = Get-FindingCountTable -Items $findings -PropertyName "code"
$findingCountsByCategory = Get-FindingCountTable -Items $findings -PropertyName "category"

$result = [pscustomobject]@{
  source = $reportInfo.Source
  reportProfileName = $resolvedValidationProfileName
  reportProfilePath = $resolvedValidationProfilePath
  passed = ($errorCount -eq 0)
  summary = [pscustomobject]@{
    charCount = $normalizedReportChars
    sectionCount = $sectionRules.Count
    foundSectionCount = @($sectionMatches | Where-Object { $_.found }).Count
    errorCount = $errorCount
    warningCount = $warningCount
    paginationRiskCount = $paginationRiskCount
    structuralIssueCount = $structuralIssueCount
    findingCountsByCode = $findingCountsByCode
    findingCountsByCategory = $findingCountsByCategory
    paginationRiskThresholds = $paginationRiskThresholds
    errorCodes = @($findings | Where-Object { $_.severity -eq "error" } | ForEach-Object { [string]$_.code } | Select-Object -Unique)
    warningCodes = @($findings | Where-Object { $_.severity -eq "warning" } | ForEach-Object { [string]$_.code } | Select-Object -Unique)
  }
  sections = $sectionMatches
  findings = $findings
}

if ($Format -eq "json") {
  $output = $result | ConvertTo-Json -Depth 8
} else {
  $linesOut = New-Object System.Collections.Generic.List[string]
  [void]$linesOut.Add("# Report Validation")
  [void]$linesOut.Add("")
  [void]$linesOut.Add("- Source: $($result.source)")
  [void]$linesOut.Add("- Passed: $($result.passed)")
  [void]$linesOut.Add("- Char count: $($result.summary.charCount)")
  [void]$linesOut.Add("- Found sections: $($result.summary.foundSectionCount)/$($result.summary.sectionCount)")
  [void]$linesOut.Add("- Errors: $($result.summary.errorCount)")
  [void]$linesOut.Add("- Warnings: $($result.summary.warningCount)")
  [void]$linesOut.Add("- Pagination risks: $($result.summary.paginationRiskCount)")
  [void]$linesOut.Add("- Pagination thresholds: long >= $($paginationRiskThresholds.longSectionChars) chars; dense >= $($paginationRiskThresholds.denseSectionChars) chars and <= $($paginationRiskThresholds.denseSectionParagraphs) paragraphs; figure cluster >= $($paginationRiskThresholds.figureClusterRefs) refs")
  [void]$linesOut.Add("- Structural issues: $($result.summary.structuralIssueCount)")
  [void]$linesOut.Add("")
  [void]$linesOut.Add("## Section Coverage")

  foreach ($sectionMatch in $sectionMatches) {
    $status = if ($sectionMatch.found) { "found" } else { "missing" }
    [void]$linesOut.Add("- $($sectionMatch.name): $status; line = $($sectionMatch.lineNumber); occurrences = $($sectionMatch.occurrenceCount); content chars = $($sectionMatch.contentChars)")
  }

  [void]$linesOut.Add("")
  [void]$linesOut.Add("## Findings")
  if ($findings.Count -eq 0) {
    [void]$linesOut.Add("- No validation findings")
  } else {
    foreach ($finding in $findings) {
      $categoryPrefix = if (-not [string]::IsNullOrWhiteSpace([string]$finding.category)) { "{0}/" -f [string]$finding.category } else { "" }
      [void]$linesOut.Add("- [$($finding.severity)] $categoryPrefix$($finding.code): $($finding.message)")
      if ($finding.PSObject.Properties.Name -contains "remediation" -and -not [string]::IsNullOrWhiteSpace([string]$finding.remediation)) {
        [void]$linesOut.Add("  Remediation: $($finding.remediation)")
      }
    }
  }

  $output = $linesOut -join [Environment]::NewLine
}

if ([string]::IsNullOrWhiteSpace($OutFile)) {
  Write-Output $output
} else {
  [System.IO.File]::WriteAllText($OutFile, $output, (New-Object System.Text.UTF8Encoding($true)))
  Write-Output "Wrote report validation to $OutFile"
}
