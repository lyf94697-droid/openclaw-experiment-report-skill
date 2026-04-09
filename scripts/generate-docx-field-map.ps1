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
$fieldMapCompositeRules = @(Get-ReportProfileFieldMapCompositeRules -Profile $reportProfile)

$labelPattern = '^(?<label>[^:\uFF1A]{1,60})[:\uFF1A]\s*(?<rest>.*)$'
$placeholderPattern = '[_\uFF3F]{2,}|\.{3,}|\uFF08\s*\uFF09|\(\s*\)|\u25A1|\u25A0'
$fullWidthColon = [string][char]0xFF1A
$choiceMarkerPattern = '[①②③④⑤⑥⑦⑧⑨⑩]'

function Normalize-OpenXmlText {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return ""
  }

  return (($Text -replace "\s+", " ").Trim())
}

function Normalize-FieldKey {
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

function Is-PlaceholderLike {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $true
  }

  return [bool]($Text -match $placeholderPattern)
}

function Get-OptionFieldInfo {
  param(
    [AllowNull()]
    [string]$Text
  )

  $normalizedText = Normalize-OpenXmlText -Text $Text
  if ([string]::IsNullOrWhiteSpace($normalizedText) -or ($normalizedText -notmatch $labelPattern)) {
    return $null
  }

  $label = $matches["label"].Trim()
  $rest = $matches["rest"]
  if ([string]::IsNullOrWhiteSpace($rest)) {
    return $null
  }

  $trimmedRest = (($rest -replace '[:\uFF1A]\s*$', '').Trim())
  if ([string]::IsNullOrWhiteSpace($trimmedRest) -or ($trimmedRest -notmatch $choiceMarkerPattern)) {
    return $null
  }

  $choiceMatches = [regex]::Matches($trimmedRest, "(?<choice>$choiceMarkerPattern\s*.*?)(?=(?:\s*$choiceMarkerPattern)|$)")
  if ($choiceMatches.Count -lt 2) {
    return $null
  }

  $choices = New-Object System.Collections.Generic.List[string]
  foreach ($choiceMatch in $choiceMatches) {
    $choiceText = Normalize-OpenXmlText -Text ($choiceMatch.Groups["choice"].Value -replace '^[√✔☑]\s*', '')
    if (-not [string]::IsNullOrWhiteSpace($choiceText)) {
      [void]$choices.Add($choiceText)
    }
  }

  if ($choices.Count -lt 2) {
    return $null
  }

  return [pscustomobject]@{
    Label = $label
    Choices = @($choices)
    HasTrailingColon = [bool]($normalizedText -match '[:\uFF1A]\s*$')
  }
}

function Resolve-OptionFieldText {
  param(
    [AllowNull()]
    [string]$TemplateText,

    [AllowNull()]
    [string]$SelectionText
  )

  $optionField = Get-OptionFieldInfo -Text $TemplateText
  if ($null -eq $optionField -or [string]::IsNullOrWhiteSpace($SelectionText)) {
    return $null
  }

  $selectionCandidates = New-Object System.Collections.Generic.List[string]
  $selectionCandidates.Add((Normalize-OpenXmlText -Text $SelectionText)) | Out-Null

  $checkedChoiceMatch = [regex]::Match($SelectionText, "[√✔☑]\s*(?<choice>$choiceMarkerPattern\s*.*?)(?=(?:\s*$choiceMarkerPattern)|$)")
  if ($checkedChoiceMatch.Success) {
    $selectionCandidates.Add((Normalize-OpenXmlText -Text $checkedChoiceMatch.Groups["choice"].Value)) | Out-Null
  }

  if ($SelectionText -match $labelPattern) {
    $selectionCandidates.Add((Normalize-OpenXmlText -Text (($matches["rest"] -replace '[:\uFF1A]\s*$', '').Trim()))) | Out-Null
  }

  $normalizedCandidates = @(
    $selectionCandidates |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    ForEach-Object { Normalize-FieldKey -Text $_ } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Select-Object -Unique
  )

  if ($normalizedCandidates.Count -eq 0) {
    return $null
  }

  $selectedIndex = -1
  for ($choiceIndex = 0; $choiceIndex -lt $optionField.Choices.Count; $choiceIndex++) {
    $normalizedChoice = Normalize-FieldKey -Text ([string]$optionField.Choices[$choiceIndex])
    foreach ($normalizedCandidate in $normalizedCandidates) {
      if (
        ($normalizedCandidate -eq $normalizedChoice) -or
        $normalizedCandidate.EndsWith($normalizedChoice) -or
        $normalizedChoice.EndsWith($normalizedCandidate)
      ) {
        $selectedIndex = $choiceIndex
        break
      }
    }

    if ($selectedIndex -ge 0) {
      break
    }
  }

  if ($selectedIndex -lt 0) {
    return $null
  }

  $renderedChoices = for ($choiceIndex = 0; $choiceIndex -lt $optionField.Choices.Count; $choiceIndex++) {
    $choiceText = Normalize-OpenXmlText -Text ([string]$optionField.Choices[$choiceIndex])
    if ($choiceIndex -eq $selectedIndex) {
      "√$choiceText"
    } else {
      $choiceText
    }
  }

  $line = ("{0}{1} {2}" -f $optionField.Label, $fullWidthColon, ($renderedChoices -join "  "))
  if ($optionField.HasTrailingColon) {
    $line += $fullWidthColon
  }

  return $line
}

function Get-ReportInput {
  param(
    [AllowNull()]
    [string]$TextPath,

    [AllowNull()]
    [string]$InlineText
  )

  if ([string]::IsNullOrWhiteSpace($TextPath) -eq [string]::IsNullOrWhiteSpace($InlineText)) {
    throw "Provide exactly one of -ReportPath or -ReportText."
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

function Add-Note {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Target,

    [Parameter(Mandatory = $true)]
    [string]$Message
  )

  $Target.Add($Message) | Out-Null
}

function Add-Diagnostic {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Diagnostics,

    [AllowNull()]
    [object]$Notes,

    [Parameter(Mandatory = $true)]
    [string]$Code,

    [Parameter(Mandatory = $true)]
    [string]$Message,

    [ValidateSet("info", "warning", "error")]
    [string]$Severity = "warning",

    [AllowNull()]
    [string]$Suggestion,

    [AllowNull()]
    [object]$Context
  )

  $diagnostic = [ordered]@{
    code = $Code
    severity = $Severity
    message = $Message
  }

  if (-not [string]::IsNullOrWhiteSpace($Suggestion)) {
    $diagnostic["suggestion"] = $Suggestion
  }

  if ($null -ne $Context) {
    $diagnostic["context"] = $Context
  }

  $Diagnostics.Add([pscustomobject]$diagnostic) | Out-Null
  if ($null -ne $Notes) {
    Add-Note -Target $Notes -Message $Message
  }
}

function Get-DiagnosticCountsByProperty {
  param(
    [AllowNull()]
    [object[]]$Diagnostics,

    [Parameter(Mandatory = $true)]
    [string]$PropertyName
  )

  $counts = [ordered]@{}
  foreach ($diagnostic in @($Diagnostics)) {
    if ($null -eq $diagnostic) {
      continue
    }

    $value = if ($diagnostic.PSObject.Properties.Name -contains $PropertyName) {
      [string]$diagnostic.$PropertyName
    } else {
      ""
    }

    if ([string]::IsNullOrWhiteSpace($value)) {
      continue
    }

    if (-not $counts.Contains($value)) {
      $counts[$value] = 0
    }

    $counts[$value]++
  }

  return $counts
}

function Test-LooksLikeSectionHeading {
  param(
    [AllowNull()]
    [string]$Text
  )

  $normalizedText = Normalize-OpenXmlText -Text $Text
  if ([string]::IsNullOrWhiteSpace($normalizedText)) {
    return $false
  }

  if ($normalizedText -match $labelPattern -or (Is-PlaceholderLike -Text $normalizedText)) {
    return $false
  }

  if ($normalizedText.Length -gt 40) {
    return $false
  }

  return [bool]($normalizedText -match '^(?:#{1,6}\s*)?(?:第?[0-9一二三四五六七八九十]+(?:章|节)?[.\)\u3001]?\s*)?.*(实验|原理|目的|内容|环境|步骤|结果|分析|总结|小结|器材|设备|任务|要求).*$')
}

function Get-CompositeCandidateSectionIds {
  param(
    [AllowNull()]
    [string]$Text
  )

  $normalizedText = Normalize-OpenXmlText -Text $Text
  if ([string]::IsNullOrWhiteSpace($normalizedText)) {
    return @()
  }

  $matchedSectionIds = New-Object System.Collections.Generic.List[string]
  foreach ($rule in $sectionRules) {
    foreach ($alias in @($rule.aliases)) {
      $aliasText = Normalize-OpenXmlText -Text ([string]$alias)
      if ([string]::IsNullOrWhiteSpace($aliasText)) {
        continue
      }

      if ($normalizedText -match [regex]::Escape($aliasText)) {
        [void]$matchedSectionIds.Add([string]$rule.id)
        break
      }
    }
  }

  return @($matchedSectionIds | Select-Object -Unique)
}

function Add-MetadataDiagnostic {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Diagnostics,

    [Parameter(Mandatory = $true)]
    [object]$Notes,

    [Parameter(Mandatory = $true)]
    [string]$Label,

    [Parameter(Mandatory = $true)]
    [string]$Location,

    [Parameter(Mandatory = $true)]
    [ValidateSet("paragraph", "table-cell")]
    [string]$Source,

    [bool]$IsOptionField = $false
  )

  $recognizedMetadataId = Resolve-KnownMetadataId -Key $Label
  if ([string]::IsNullOrWhiteSpace($recognizedMetadataId)) {
    Add-Diagnostic `
      -Diagnostics $Diagnostics `
      -Notes $Notes `
      -Code "unrecognized_template_metadata_label" `
      -Severity "warning" `
      -Message ("Template label was detected but the profile does not recognize it: {0}" -f $Label) `
      -Suggestion "Add a matching metadata alias to metadataFields in the report profile, or rename the template label to an existing metadata field." `
      -Context ([ordered]@{
          label = $Label
          location = $Location
          source = $Source
          optionField = [bool]$IsOptionField
        })
    return
  }

  Add-Diagnostic `
    -Diagnostics $Diagnostics `
    -Notes $Notes `
    -Code "missing_metadata_value" `
    -Severity "warning" `
    -Message ("Template label was detected but no value was available: {0}" -f $Label) `
    -Suggestion "Provide this value in metadata JSON, or include the same metadata label and value in the report header." `
    -Context ([ordered]@{
        label = $Label
        metadataId = $recognizedMetadataId
        location = $Location
        source = $Source
        optionField = [bool]$IsOptionField
      })
}

function New-MetadataRule {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Id,

    [Parameter(Mandatory = $true)]
    [string[]]$Aliases
  )

  return [pscustomobject]@{
    id = $Id
    aliases = $Aliases
  }
}

function New-SectionRule {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Id,

    [Parameter(Mandatory = $true)]
    [string[]]$Aliases
  )

  return [pscustomobject]@{
    id = $Id
    aliases = $Aliases
  }
}

function Get-ProfileSectionFieldMapId {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$SectionField
  )

  if ($SectionField.PSObject.Properties.Name -contains "fieldMapId" -and -not [string]::IsNullOrWhiteSpace([string]$SectionField.fieldMapId)) {
    return [string]$SectionField.fieldMapId
  }

  $sectionKey = if ($SectionField.PSObject.Properties.Name -contains "key") { [string]$SectionField.key } else { "" }
  $resolvedSectionId = Get-ReportProfileSectionIdFromKey -Key $sectionKey
  if (-not [string]::IsNullOrWhiteSpace($resolvedSectionId)) {
    return $resolvedSectionId
  }

  throw "Profile section field is missing both fieldMapId and key."
}

$metadataRules = @(
  (New-MetadataRule -Id "course_name" -Aliases @("课程名称", "课程名", "课程", "course", "coursename")),
  (New-MetadataRule -Id "experiment_name" -Aliases @("实验名称", "实验名", "实验题目", "题目", "experiment", "experimentname", "experimenttitle", "labreporttitle")),
  (New-MetadataRule -Id "experiment_property" -Aliases @("实验性质", "实验类型", "experimentproperty", "experimenttype", "labtype")),
  (New-MetadataRule -Id "student_name" -Aliases @("姓名", "name", "studentname")),
  (New-MetadataRule -Id "student_id" -Aliases @("学号", "studentid", "studentnumber", "id")),
  (New-MetadataRule -Id "class_name" -Aliases @("班级", "classname", "class")),
  (New-MetadataRule -Id "teacher_name" -Aliases @("指导教师", "教师", "老师", "teacher", "teachername")),
  (New-MetadataRule -Id "experiment_date" -Aliases @("实验时间", "实验日期", "实验日期时间", "experimentdate", "experimenttime")),
  (New-MetadataRule -Id "experiment_location" -Aliases @("实验地点", "实验位置", "experimentlocation")),
  (New-MetadataRule -Id "date" -Aliases @("日期", "实验日期", "date")),
  (New-MetadataRule -Id "college" -Aliases @("学院", "college")),
  (New-MetadataRule -Id "major" -Aliases @("专业", "major"))
)

$sectionRules = @(
  foreach ($sectionField in @($reportProfile.sectionFields)) {
    $aliases = @($sectionField.aliases)
    if ($aliases.Count -eq 0 -and $sectionField.PSObject.Properties.Name -contains "heading") {
      $aliases = @([string]$sectionField.heading)
    }

    New-SectionRule -Id (Get-ProfileSectionFieldMapId -SectionField $sectionField) -Aliases $aliases
  }
)

$metadataAliasLookup = @{}
foreach ($rule in $metadataRules) {
  foreach ($alias in $rule.aliases) {
    $normalizedAlias = Normalize-FieldKey -Text $alias
    if (-not [string]::IsNullOrWhiteSpace($normalizedAlias)) {
      $metadataAliasLookup[$normalizedAlias] = $rule.id
    }
  }
}

function Resolve-KnownMetadataId {
  param(
    [AllowNull()]
    [string]$Key
  )

  $normalized = Normalize-FieldKey -Text $Key
  if ([string]::IsNullOrWhiteSpace($normalized)) {
    return $null
  }

  if ($metadataAliasLookup.ContainsKey($normalized)) {
    return [string]$metadataAliasLookup[$normalized]
  }

  return $null
}

function Resolve-MetadataId {
  param(
    [AllowNull()]
    [string]$Key
  )

  $normalized = Normalize-FieldKey -Text $Key
  if ([string]::IsNullOrWhiteSpace($normalized)) {
    return $null
  }

  if ($metadataAliasLookup.ContainsKey($normalized)) {
    return $metadataAliasLookup[$normalized]
  }

  return $normalized
}

function Add-MetadataValue {
  param(
    [Parameter(Mandatory = $true)]
    [hashtable]$MetadataValues,

    [AllowNull()]
    [string]$Key,

    [AllowNull()]
    [string]$Value
  )

  if ([string]::IsNullOrWhiteSpace($Key) -or [string]::IsNullOrWhiteSpace($Value)) {
    return
  }

  $resolvedId = Resolve-MetadataId -Key $Key
  if ([string]::IsNullOrWhiteSpace($resolvedId)) {
    return
  }

  $MetadataValues[$resolvedId] = $Value.Trim()
}

function Get-MetadataValue {
  param(
    [Parameter(Mandatory = $true)]
    [hashtable]$MetadataValues,

    [AllowNull()]
    [string]$Key
  )

  $resolvedId = Resolve-MetadataId -Key $Key
  if ([string]::IsNullOrWhiteSpace($resolvedId)) {
    return $null
  }

  if ($MetadataValues.ContainsKey($resolvedId)) {
    return [string]$MetadataValues[$resolvedId]
  }

  return $null
}

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

function Import-OptionalMetadata {
  param(
    [AllowNull()]
    [string]$PathToJson,

    [AllowNull()]
    [string]$InlineJson,

    [Parameter(Mandatory = $true)]
    [hashtable]$MetadataValues
  )

  if ([string]::IsNullOrWhiteSpace($PathToJson) -and [string]::IsNullOrWhiteSpace($InlineJson)) {
    return
  }

  if ([string]::IsNullOrWhiteSpace($PathToJson) -eq [string]::IsNullOrWhiteSpace($InlineJson)) {
    throw "Provide zero or one of -MetadataPath and -MetadataJson."
  }

  if (-not [string]::IsNullOrWhiteSpace($PathToJson)) {
    $resolvedPath = Resolve-Path -LiteralPath $PathToJson
    $metadataRoot = (Get-Content -LiteralPath $resolvedPath.Path -Raw -Encoding UTF8) | ConvertFrom-Json
  } else {
    $metadataRoot = $InlineJson | ConvertFrom-Json
  }

  if ($null -eq $metadataRoot) {
    return
  }

  $metadataTable = ConvertTo-PlainHashtable -InputObject $metadataRoot
  if ($metadataTable.ContainsKey("metadata")) {
    $metadataTable = ConvertTo-PlainHashtable -InputObject $metadataTable["metadata"]
  }

  foreach ($entry in $metadataTable.GetEnumerator()) {
    if ($null -eq $entry.Value) {
      continue
    }

    $value = Normalize-OpenXmlText -Text ([string]$entry.Value)
    Add-MetadataValue -MetadataValues $MetadataValues -Key ([string]$entry.Key) -Value $value
  }
}

function Get-HeadingPattern {
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$Aliases
  )

  $escapedAliases = foreach ($alias in $Aliases) {
    [regex]::Escape($alias)
  }

  return '^(?:#{1,6}\s*)?(?:第?[0-9一二三四五六七八九十]+(?:章|节)?[.\)\u3001]?\s*)?(?:' + ($escapedAliases -join '|') + ')\s*(?:[:\uFF1A])?$'
}

function Resolve-SectionRule {
  param(
    [AllowNull()]
    [string]$HeadingText
  )

  if ([string]::IsNullOrWhiteSpace($HeadingText)) {
    return $null
  }

  foreach ($rule in $sectionRules) {
    if ($HeadingText -match (Get-HeadingPattern -Aliases $rule.aliases)) {
      return $rule
    }
  }

  return $null
}

function Convert-ToParagraphList {
  param(
    [AllowNull()]
    [string]$Text
  )

  $paragraphs = New-Object System.Collections.Generic.List[string]
  if ([string]::IsNullOrWhiteSpace($Text)) {
    return @()
  }

  $buffer = New-Object System.Collections.Generic.List[string]
  foreach ($line in ($Text -split '\r?\n')) {
    $trimmedLine = $line.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmedLine)) {
      if ($buffer.Count -gt 0) {
        [void]$paragraphs.Add((Normalize-OpenXmlText -Text ($buffer -join ' ')))
        $buffer.Clear()
      }
      continue
    }

    $startsStructuredParagraph = $trimmedLine -match '^(?:[-*•]\s+|\d+[.\)\u3001]\s*|[（(]?\d+[）)]\s*|[一二三四五六七八九十]+[.\u3001]\s*|(?:主机|步骤|结果|说明|分析)\s*[A-Z0-9一二三四五六七八九十]?\s*[:\uFF1A])'
    if ($startsStructuredParagraph -and $buffer.Count -gt 0) {
      [void]$paragraphs.Add((Normalize-OpenXmlText -Text ($buffer -join ' ')))
      $buffer.Clear()
    }

    [void]$buffer.Add($trimmedLine)
  }

  if ($buffer.Count -gt 0) {
    [void]$paragraphs.Add((Normalize-OpenXmlText -Text ($buffer -join ' ')))
  }

  $filteredParagraphs = New-Object System.Collections.Generic.List[string]
  foreach ($paragraph in $paragraphs) {
    if (-not [string]::IsNullOrWhiteSpace($paragraph)) {
      [void]$filteredParagraphs.Add($paragraph)
    }
  }
  $paragraphs = $filteredParagraphs

  if ($paragraphs.Count -eq 0) {
    foreach ($paragraph in ($Text -split '(?:\r?\n){2,}')) {
      $normalizedParagraph = Normalize-OpenXmlText -Text $paragraph
      if (-not [string]::IsNullOrWhiteSpace($normalizedParagraph)) {
        [void]$paragraphs.Add($normalizedParagraph)
      }
    }
  }

  if ($paragraphs.Count -eq 0) {
    $singleParagraph = Normalize-OpenXmlText -Text $Text
    if (-not [string]::IsNullOrWhiteSpace($singleParagraph)) {
      [void]$paragraphs.Add($singleParagraph)
    }
  }

  return $paragraphs.ToArray()
}

function Get-ReportAnalysis {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Text,

    [Parameter(Mandatory = $true)]
    [hashtable]$MetadataValues
  )

  $lines = @($Text -split "\r?\n")
  $trimmedLines = @(foreach ($line in $lines) { $line.Trim() })

  foreach ($line in $trimmedLines) {
    if ($line -match $labelPattern) {
      $label = $matches["label"].Trim()
      $rest = Normalize-OpenXmlText -Text $matches["rest"]
      if (-not [string]::IsNullOrWhiteSpace($rest) -and -not (Is-PlaceholderLike -Text $rest)) {
        Add-MetadataValue -MetadataValues $MetadataValues -Key $label -Value $rest
      }
    }
  }

  $sectionMatches = New-Object System.Collections.Generic.List[object]
  foreach ($rule in $sectionRules) {
    $headingIndex = $null
    $headingText = $null
    $headingPattern = Get-HeadingPattern -Aliases $rule.aliases
    for ($lineIndex = 0; $lineIndex -lt $trimmedLines.Count; $lineIndex++) {
      $lineText = $trimmedLines[$lineIndex]
      if ([string]::IsNullOrWhiteSpace($lineText)) {
        continue
      }

      if ($lineText -match $headingPattern) {
        $headingIndex = $lineIndex
        $headingText = $lineText
        break
      }
    }

    if ($null -ne $headingIndex) {
      $sectionMatches.Add([pscustomobject]@{
          id = $rule.id
          headingIndex = $headingIndex
          headingText = $headingText
        }) | Out-Null
    }
  }

  $sectionsById = @{}
  $orderedMatches = @($sectionMatches | Sort-Object headingIndex)
  for ($i = 0; $i -lt $orderedMatches.Count; $i++) {
    $current = $orderedMatches[$i]
    $nextHeadingIndex = if (($i + 1) -lt $orderedMatches.Count) { $orderedMatches[$i + 1].headingIndex } else { $lines.Count }
    $contentLines = if (($current.headingIndex + 1) -lt $nextHeadingIndex) { $lines[($current.headingIndex + 1)..($nextHeadingIndex - 1)] } else { @() }
    $contentText = ($contentLines -join [Environment]::NewLine).Trim()
    $sectionsById[$current.id] = [pscustomobject]@{
      headingText = $current.headingText
      text = $contentText
      paragraphs = @(Convert-ToParagraphList -Text $contentText)
    }
  }

  return @{
    SectionsById = $sectionsById
  }
}

function Get-TemplateAnalysis {
  param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath
  )

  $analysisText = & (Join-Path $repoRoot "scripts\extract-docx-template.ps1") -Path $DocxPath -Format json | Out-String
  $analysis = $analysisText | ConvertFrom-Json
  if ($null -eq $analysis) {
    throw "Failed to parse template analysis JSON."
  }

  return $analysis
}

function Add-SectionParagraphBlock {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Target,

    [Parameter(Mandatory = $true)]
    [string]$Heading,

    [AllowNull()]
    [object]$SectionInfo
  )

  if ($null -eq $SectionInfo -or @($SectionInfo.paragraphs).Count -eq 0) {
    return $false
  }

  [void]$Target.Add($Heading)
  foreach ($paragraph in @($SectionInfo.paragraphs)) {
    if (-not [string]::IsNullOrWhiteSpace([string]$paragraph)) {
      [void]$Target.Add([string]$paragraph)
    }
  }

  return $true
}

function Get-CompositeCellFillSpec {
  param(
    [AllowNull()]
    [string]$CellText,

    [Parameter(Mandatory = $true)]
    [hashtable]$SectionsById
  )

  $normalizedCellText = Normalize-OpenXmlText -Text $CellText
  if ([string]::IsNullOrWhiteSpace($normalizedCellText)) {
    return $null
  }

  $paragraphs = New-Object System.Collections.Generic.List[string]
  $mappedSectionIds = New-Object System.Collections.Generic.List[string]

  foreach ($compositeRule in $fieldMapCompositeRules) {
    $matchesRule = $true
    foreach ($pattern in @($compositeRule.matchAll)) {
      if ([string]::IsNullOrWhiteSpace([string]$pattern) -or $normalizedCellText -notmatch [string]$pattern) {
        $matchesRule = $false
        break
      }
    }

    if (-not $matchesRule) {
      continue
    }

    foreach ($block in @($compositeRule.blocks)) {
      $blockAdded = $false
      foreach ($sectionId in @($block.sectionIds)) {
        if (-not $SectionsById.ContainsKey([string]$sectionId)) {
          continue
        }

        $sectionInfo = $SectionsById[[string]$sectionId]
        if ($null -eq $sectionInfo -or @($sectionInfo.paragraphs).Count -eq 0) {
          continue
        }

        if (-not $blockAdded -and -not [string]::IsNullOrWhiteSpace([string]$block.heading)) {
          [void]$paragraphs.Add([string]$block.heading)
          $blockAdded = $true
        }

        foreach ($paragraph in @($sectionInfo.paragraphs)) {
          if (-not [string]::IsNullOrWhiteSpace([string]$paragraph)) {
            [void]$paragraphs.Add([string]$paragraph)
          }
        }
        [void]$mappedSectionIds.Add([string]$sectionId)
      }
    }

    if ($paragraphs.Count -gt 0 -and $mappedSectionIds.Count -gt 0) {
      return [pscustomobject]@{
        Paragraphs = @($paragraphs)
        MappedSectionIds = @($mappedSectionIds | Select-Object -Unique)
      }
    }
  }

  return $null
}

function Get-CompositeTableBodyFillSpec {
  param(
    [AllowNull()]
    [object[]]$CompositeEntries,

    [Parameter(Mandatory = $true)]
    [object]$TableBlock
  )

  $entries = @($CompositeEntries)
  if ($entries.Count -eq 0) {
    return $null
  }

  $orderedEntries = @($entries | Sort-Object RowIndex, CellIndex)
  $startRowIndex = [int]$orderedEntries[0].RowIndex
  $throughRowIndex = [int]$orderedEntries[-1].RowIndex
  $tableRowCount = @($TableBlock.rows).Count

  if ($startRowIndex -le 1 -or $throughRowIndex -gt $tableRowCount) {
    return $null
  }

  $entryRowIndices = @($orderedEntries | ForEach-Object { [int]$_.RowIndex } | Sort-Object -Unique)
  for ($rowIndex = $startRowIndex; $rowIndex -le $throughRowIndex; $rowIndex++) {
    if ($rowIndex -notin $entryRowIndices) {
      return $null
    }
  }

  $paragraphs = New-Object System.Collections.Generic.List[string]
  $mappedSectionIds = New-Object System.Collections.Generic.List[string]

  foreach ($entry in $orderedEntries) {
    foreach ($paragraph in @($entry.Paragraphs)) {
      if (-not [string]::IsNullOrWhiteSpace([string]$paragraph)) {
        [void]$paragraphs.Add([string]$paragraph)
      }
    }

    foreach ($sectionId in @($entry.MappedSectionIds)) {
      if (-not [string]::IsNullOrWhiteSpace([string]$sectionId)) {
        [void]$mappedSectionIds.Add([string]$sectionId)
      }
    }
  }

  if ($paragraphs.Count -eq 0 -or $mappedSectionIds.Count -eq 0) {
    return $null
  }

  return [pscustomobject]@{
    StartLocation = [string]$orderedEntries[0].Location
    ThroughLocation = [string]$orderedEntries[-1].Location
    Paragraphs = @($paragraphs)
    MappedSectionIds = @($mappedSectionIds | Select-Object -Unique)
  }
}

function ConvertTo-JsonComparable {
  param(
    [AllowNull()]
    [object]$Value
  )

  if ($null -eq $Value) {
    return "null"
  }

  return ($Value | ConvertTo-Json -Depth 8 -Compress)
}

function Add-FieldMapEntry {
  param(
    [Parameter(Mandatory = $true)]
    [System.Collections.IDictionary]$FieldMap,

    [Parameter(Mandatory = $true)]
    [string]$Key,

    [Parameter(Mandatory = $true)]
    [object]$Value,

    [Parameter(Mandatory = $true)]
    [object]$Notes,

    [AllowNull()]
    [object]$Diagnostics
  )

  if ([string]::IsNullOrWhiteSpace($Key)) {
    return $false
  }

  if ($FieldMap.Contains($Key)) {
    if ((ConvertTo-JsonComparable -Value $FieldMap[$Key]) -ne (ConvertTo-JsonComparable -Value $Value)) {
      Add-Diagnostic `
        -Diagnostics $Diagnostics `
        -Notes $Notes `
        -Code "duplicate_template_field_conflict" `
        -Severity "warning" `
        -Message ("Duplicate template field skipped because a different value was already selected: {0}" -f $Key) `
        -Suggestion "Make the target placeholder unique in the template, or remove the conflicting duplicate block." `
        -Context ([ordered]@{
            key = $Key
          })
    }
    return $false
  }

  $FieldMap[$Key] = $Value
  return $true
}

$resolvedTemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path
$reportInfo = Get-ReportInput -TextPath $ReportPath -InlineText $ReportText
$metadataValues = @{}
Import-OptionalMetadata -PathToJson $MetadataPath -InlineJson $MetadataJson -MetadataValues $metadataValues
$reportAnalysis = Get-ReportAnalysis -Text ([string]$reportInfo.Text) -MetadataValues $metadataValues
$templateAnalysis = Get-TemplateAnalysis -DocxPath $resolvedTemplatePath

$fieldMap = [ordered]@{}
$notes = New-Object System.Collections.Generic.List[string]
$diagnostics = New-Object System.Collections.Generic.List[object]
$mappedMetadataCount = 0
$mappedSectionCount = 0

$blocks = @($templateAnalysis.blocks)
for ($blockIndex = 0; $blockIndex -lt $blocks.Count; $blockIndex++) {
  $block = $blocks[$blockIndex]
  if ($block.type -eq "paragraph") {
    $paragraphText = Normalize-OpenXmlText -Text ([string]$block.text)
    if ([string]::IsNullOrWhiteSpace($paragraphText)) {
      continue
    }

    $sectionRule = Resolve-SectionRule -HeadingText $paragraphText
    if ($null -ne $sectionRule) {
      if ($reportAnalysis.SectionsById.ContainsKey($sectionRule.id)) {
        $sectionInfo = $reportAnalysis.SectionsById[$sectionRule.id]
        if (@($sectionInfo.paragraphs).Count -gt 0) {
          $nextBlock = if (($blockIndex + 1) -lt $blocks.Count) { $blocks[$blockIndex + 1] } else { $null }
          $useAfter = ($null -ne $nextBlock) -and $nextBlock.type -eq "paragraph" -and (Is-PlaceholderLike -Text ([string]$nextBlock.text))
          $sectionValue = if ($useAfter) {
            [ordered]@{
              mode = "after"
              paragraphs = @($sectionInfo.paragraphs)
            }
          } elseif (@($sectionInfo.paragraphs).Count -eq 1) {
            $sectionInfo.paragraphs[0]
          } else {
            @($sectionInfo.paragraphs)
          }

          if (Add-FieldMapEntry -FieldMap $fieldMap -Key $paragraphText -Value $sectionValue -Notes $notes -Diagnostics $diagnostics) {
            $mappedSectionCount++
          }
        } else {
          Add-Diagnostic `
            -Diagnostics $diagnostics `
            -Notes $notes `
            -Code "empty_report_section" `
            -Severity "warning" `
            -Message ("Report section is present but empty: {0}" -f $paragraphText) `
            -Suggestion "Fill this section in the report draft, or remove the placeholder section from the template if it is optional." `
            -Context ([ordered]@{
                heading = $paragraphText
                sectionId = [string]$sectionRule.id
                location = [string]$block.location
              })
        }
      } else {
        Add-Diagnostic `
          -Diagnostics $diagnostics `
          -Notes $notes `
          -Code "missing_report_section" `
          -Severity "warning" `
          -Message ("Template section has no matching report section: {0}" -f $paragraphText) `
          -Suggestion "Add this section to the report draft, or extend sectionFields aliases in the report profile if the template uses a different heading name." `
          -Context ([ordered]@{
              heading = $paragraphText
              sectionId = [string]$sectionRule.id
              location = [string]$block.location
            })
      }

      continue
    }

    if (Test-LooksLikeSectionHeading -Text $paragraphText) {
      Add-Diagnostic `
        -Diagnostics $diagnostics `
        -Notes $notes `
        -Code "unrecognized_template_section_heading" `
        -Severity "warning" `
        -Message ("Template heading looks like a report section, but no profile section rule matched it: {0}" -f $paragraphText) `
        -Suggestion "Add this heading as a sectionFields alias in the report profile, or rename the template heading to an existing section heading." `
        -Context ([ordered]@{
            heading = $paragraphText
            location = [string]$block.location
          })
    }

    if ($paragraphText -match $labelPattern) {
      $label = $matches["label"].Trim()
      $rest = $matches["rest"]
      $optionFieldValue = Get-MetadataValue -MetadataValues $metadataValues -Key $label
      $resolvedOptionFieldText = Resolve-OptionFieldText -TemplateText $paragraphText -SelectionText $optionFieldValue
      if (-not [string]::IsNullOrWhiteSpace($resolvedOptionFieldText)) {
        if (Add-FieldMapEntry -FieldMap $fieldMap -Key ([string]$block.location) -Value $resolvedOptionFieldText -Notes $notes -Diagnostics $diagnostics) {
          $mappedMetadataCount++
        }
        continue
      }

      if ($null -ne (Get-OptionFieldInfo -Text $paragraphText)) {
        Add-MetadataDiagnostic -Diagnostics $diagnostics -Notes $notes -Label $label -Location ([string]$block.location) -Source "paragraph" -IsOptionField $true
        continue
      }

      if ([string]::IsNullOrWhiteSpace($rest) -or (Is-PlaceholderLike -Text $rest)) {
        $value = Get-MetadataValue -MetadataValues $metadataValues -Key $label
        if (-not [string]::IsNullOrWhiteSpace($value)) {
          if (Add-FieldMapEntry -FieldMap $fieldMap -Key $label -Value $value -Notes $notes -Diagnostics $diagnostics) {
            $mappedMetadataCount++
          }
        } else {
          Add-MetadataDiagnostic -Diagnostics $diagnostics -Notes $notes -Label $label -Location ([string]$block.location) -Source "paragraph"
        }
      }
    }

    continue
  }

  if ($block.type -eq "table") {
    $tableCompositeEntries = New-Object System.Collections.Generic.List[object]

    foreach ($row in @($block.rows)) {
      $cells = @($row.cells)
      for ($cellIndex = 0; $cellIndex -lt $cells.Count; $cellIndex++) {
        $cell = $cells[$cellIndex]
        $cellText = Normalize-OpenXmlText -Text ([string]$cell.text)
        if ([string]::IsNullOrWhiteSpace($cellText)) {
          continue
        }

        $compositeCellFill = Get-CompositeCellFillSpec -CellText $cellText -SectionsById $reportAnalysis.SectionsById
        if ($null -ne $compositeCellFill) {
          $tableCompositeEntries.Add([pscustomobject]@{
              Location = [string]$cell.location
              RowIndex = [int]$row.row
              CellIndex = $cellIndex + 1
              Paragraphs = @($compositeCellFill.Paragraphs)
              MappedSectionIds = @($compositeCellFill.MappedSectionIds)
            }) | Out-Null
          continue
        }

        $compositeCandidateSectionIds = @(Get-CompositeCandidateSectionIds -Text $cellText)
        if ($compositeCandidateSectionIds.Count -ge 2) {
          Add-Diagnostic `
            -Diagnostics $diagnostics `
            -Notes $notes `
            -Code "unmatched_composite_template_cell" `
            -Severity "warning" `
            -Message ("Template cell looks like a composite body container, but no profile composite rule matched it: {0}" -f $cellText) `
            -Suggestion "Add a matching fieldMapCompositeRules entry to the report profile, or split this cell into separate section placeholders." `
            -Context ([ordered]@{
                cellText = $cellText
                location = [string]$cell.location
                matchedSectionIds = @($compositeCandidateSectionIds)
              })
        }

        if ($cellText -match $labelPattern) {
          $label = $matches["label"].Trim()
          $rest = $matches["rest"]
          $optionFieldValue = Get-MetadataValue -MetadataValues $metadataValues -Key $label
          $resolvedOptionFieldText = Resolve-OptionFieldText -TemplateText $cellText -SelectionText $optionFieldValue
          if (-not [string]::IsNullOrWhiteSpace($resolvedOptionFieldText)) {
            if (Add-FieldMapEntry -FieldMap $fieldMap -Key ([string]$cell.location) -Value $resolvedOptionFieldText -Notes $notes -Diagnostics $diagnostics) {
              $mappedMetadataCount++
            }
            continue
          }

          if ($null -ne (Get-OptionFieldInfo -Text $cellText)) {
            Add-MetadataDiagnostic -Diagnostics $diagnostics -Notes $notes -Label $label -Location ([string]$cell.location) -Source "table-cell" -IsOptionField $true
            continue
          }

          if ([string]::IsNullOrWhiteSpace($rest) -or (Is-PlaceholderLike -Text $rest)) {
            $value = Get-MetadataValue -MetadataValues $metadataValues -Key $label
            if (-not [string]::IsNullOrWhiteSpace($value)) {
              if (Add-FieldMapEntry -FieldMap $fieldMap -Key $label -Value $value -Notes $notes -Diagnostics $diagnostics) {
                $mappedMetadataCount++
              }
            } else {
              Add-MetadataDiagnostic -Diagnostics $diagnostics -Notes $notes -Label $label -Location ([string]$cell.location) -Source "table-cell"
            }
          }
          continue
        }

        if (($cellIndex + 1) -lt $cells.Count) {
          $nextCell = $cells[$cellIndex + 1]
          $nextText = Normalize-OpenXmlText -Text ([string]$nextCell.text)
          if (Is-PlaceholderLike -Text $nextText) {
            $value = Get-MetadataValue -MetadataValues $metadataValues -Key $cellText
            if (-not [string]::IsNullOrWhiteSpace($value)) {
              if (Add-FieldMapEntry -FieldMap $fieldMap -Key $cellText -Value $value -Notes $notes -Diagnostics $diagnostics) {
                $mappedMetadataCount++
              }
            } else {
              Add-MetadataDiagnostic -Diagnostics $diagnostics -Notes $notes -Label $cellText -Location ([string]$nextCell.location) -Source "table-cell"
            }
          } elseif (Test-LooksLikeSectionHeading -Text $cellText) {
            Add-Diagnostic `
              -Diagnostics $diagnostics `
              -Notes $notes `
              -Code "unrecognized_template_section_heading" `
              -Severity "warning" `
              -Message ("Template heading looks like a report section, but no profile section rule matched it: {0}" -f $cellText) `
              -Suggestion "Add this heading as a sectionFields alias in the report profile, or rename the template heading to an existing section heading." `
              -Context ([ordered]@{
                  heading = $cellText
                  location = [string]$cell.location
                })
          }
        } elseif (Test-LooksLikeSectionHeading -Text $cellText) {
          Add-Diagnostic `
            -Diagnostics $diagnostics `
            -Notes $notes `
            -Code "unrecognized_template_section_heading" `
            -Severity "warning" `
            -Message ("Template heading looks like a report section, but no profile section rule matched it: {0}" -f $cellText) `
            -Suggestion "Add this heading as a sectionFields alias in the report profile, or rename the template heading to an existing section heading." `
            -Context ([ordered]@{
                heading = $cellText
                location = [string]$cell.location
              })
        }
      }
    }

    $tableCompositeBodyFill = $null
    if ($tableCompositeEntries.Count -gt 0) {
      $tableCompositeBodyFill = Get-CompositeTableBodyFillSpec -CompositeEntries $tableCompositeEntries.ToArray() -TableBlock $block
    }
    if ($null -ne $tableCompositeBodyFill) {
      $tableBodyValue = [ordered]@{
        mode = "after-table"
        through = $tableCompositeBodyFill.ThroughLocation
        paragraphs = @($tableCompositeBodyFill.Paragraphs)
      }

      if (Add-FieldMapEntry -FieldMap $fieldMap -Key $tableCompositeBodyFill.StartLocation -Value $tableBodyValue -Notes $notes -Diagnostics $diagnostics) {
        $mappedSectionCount += @($tableCompositeBodyFill.MappedSectionIds).Count
        Add-Diagnostic `
          -Diagnostics $diagnostics `
          -Notes $notes `
          -Code "composite_body_after_table" `
          -Severity "info" `
          -Message ("Composite table body will be moved after the cover table: {0} -> {1}" -f $tableCompositeBodyFill.StartLocation, $tableCompositeBodyFill.ThroughLocation) `
          -Suggestion "This is expected for cover/body hybrid templates. If the placement is wrong, adjust the template rows or fieldMapCompositeRules." `
          -Context ([ordered]@{
              startLocation = [string]$tableCompositeBodyFill.StartLocation
              throughLocation = [string]$tableCompositeBodyFill.ThroughLocation
              mappedSectionIds = @($tableCompositeBodyFill.MappedSectionIds)
            })
      }
    } else {
      foreach ($compositeEntry in $tableCompositeEntries) {
        if (Add-FieldMapEntry -FieldMap $fieldMap -Key $compositeEntry.Location -Value @($compositeEntry.Paragraphs) -Notes $notes -Diagnostics $diagnostics) {
          $mappedSectionCount += @($compositeEntry.MappedSectionIds).Count
        }
      }
    }
  }
}

$result = [ordered]@{
  templatePath = $resolvedTemplatePath
  reportSource = $reportInfo.Source
  reportProfileName = [string]$reportProfile.name
  reportProfilePath = [string]$reportProfile.resolvedProfilePath
  summary = [ordered]@{
    metadataValueCount = $metadataValues.Count
    reportSectionCount = $reportAnalysis.SectionsById.Count
    fieldCount = $fieldMap.Count
    mappedMetadataCount = $mappedMetadataCount
    mappedSectionCount = $mappedSectionCount
    diagnosticCount = $diagnostics.Count
    diagnosticCountsByCode = (Get-DiagnosticCountsByProperty -Diagnostics ([object[]]$diagnostics) -PropertyName "code")
    diagnosticCountsBySeverity = (Get-DiagnosticCountsByProperty -Diagnostics ([object[]]$diagnostics) -PropertyName "severity")
    noteCount = $notes.Count
  }
  fieldMap = $fieldMap
  diagnostics = ([object[]]$diagnostics)
  notes = @($notes)
}

if ($Format -eq "json") {
  $output = $result | ConvertTo-Json -Depth 10
} else {
  $lines = New-Object System.Collections.Generic.List[string]
  [void]$lines.Add("# DOCX Field Map")
  [void]$lines.Add("")
  [void]$lines.Add("- Template: $($result.templatePath)")
  [void]$lines.Add("- Report source: $($result.reportSource)")
  [void]$lines.Add("- Metadata values: $($result.summary.metadataValueCount)")
  [void]$lines.Add("- Report sections: $($result.summary.reportSectionCount)")
  [void]$lines.Add("- Generated fields: $($result.summary.fieldCount)")
  [void]$lines.Add("- Diagnostics: $($result.summary.diagnosticCount)")
  [void]$lines.Add("- Notes: $($result.summary.noteCount)")
  [void]$lines.Add("")
  [void]$lines.Add("## Field Map JSON")
  [void]$lines.Add('```json')
  $fieldMapJson = $result.fieldMap | ConvertTo-Json -Depth 10
  [void]$lines.Add($fieldMapJson)
  [void]$lines.Add('```')
  [void]$lines.Add("")
  [void]$lines.Add("## Diagnostics")
  if ($diagnostics.Count -eq 0) {
    [void]$lines.Add("- None")
  } else {
    foreach ($diagnostic in $diagnostics) {
      $line = "- [$($diagnostic.severity)] [$($diagnostic.code)] $($diagnostic.message)"
      if ($diagnostic.PSObject.Properties.Name -contains "suggestion" -and -not [string]::IsNullOrWhiteSpace([string]$diagnostic.suggestion)) {
        $line = "$line Suggestion: $([string]$diagnostic.suggestion)"
      }
      [void]$lines.Add($line)
    }
  }
  [void]$lines.Add("")
  [void]$lines.Add("## Notes")
  if ($notes.Count -eq 0) {
    [void]$lines.Add("- None")
  } else {
    foreach ($note in $notes) {
      [void]$lines.Add("- $note")
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
  Write-Output "Wrote field map to $OutFile"
}
